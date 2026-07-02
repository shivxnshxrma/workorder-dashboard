export interface MerlinConfig {
  apiUrl: string;
  username: string;
  password: string;
}

export interface SiteLocation {
  id: string;
  name: string;
  code?: string;
}

export class MerlinClient {
  private apiUrl: string;
  private username: string;
  private password: string;
  private token: string = '';
  private onLog: (msg: string) => void;

  private cache: {
    clients: Record<string, string>;
    locations: Record<string, string>;
    forms: Record<string, string>;
    work_order_types: Record<string, string>;
    teams: Record<string, string>;
    tags: Record<string, string>;
    assets: Record<string, string>;
    users: Record<string, string>;
  } = {
    clients: {},
    locations: {},
    forms: {},
    work_order_types: {},
    teams: {},
    tags: {},
    assets: {},
    users: {},
  };

  constructor(config: MerlinConfig, onLog: (msg: string) => void) {
    // Standardize URL to end with slash
    this.apiUrl = config.apiUrl.endsWith('/') ? config.apiUrl : `${config.apiUrl}/`;
    this.username = config.username;
    this.password = config.password;
    this.onLog = onLog;
  }

  private log(message: string) {
    this.onLog(message);
  }

  private async requestJson(method: string, endpoint: string, payload: any = null, params: Record<string, any> = {}) {
    let url = `${this.apiUrl}${endpoint}`;
    if (url.endsWith('//')) {
      url = url.slice(0, -1);
    }

    const queryParams = new URLSearchParams();
    Object.entries(params).forEach(([key, val]) => {
      if (val !== undefined && val !== null) {
        queryParams.append(key, String(val));
      }
    });

    const queryString = queryParams.toString();
    const targetUrl = queryString ? `${url}?${queryString}` : url;

    const headers: Record<string, string> = {};
    if (this.token) {
      headers['Authorization'] = `Bearer ${this.token}`;
    }

    try {
      const response = await fetch('/api/proxy', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          url: targetUrl,
          method,
          headers,
          body: payload,
        }),
      });

      if (!response.ok) {
        return { status: response.status, data: { error: `HTTP error ${response.status}` } };
      }

      const result = await response.json();
      return { status: result.status, data: result.data };
    } catch (err: any) {
      return { status: 500, data: { error: err.message || 'Network error' } };
    }
  }

  public async authenticate(): Promise<boolean> {
    this.log(`Authenticating user: ${this.username}...`);
    const loginUrl = 'auth/login/email/';
    const payload = {
      email: this.username,
      password: this.password,
      app_id: 'dashboard',
    };

    const { status, data } = await this.requestJson('POST', loginUrl, payload);
    if (status === 200) {
      const token =
        data?.access_token ||
        data?.token ||
        data?.data?.access_token ||
        data?.data?.token;

      if (token) {
        this.token = token;
        this.log('Authentication successful.');
        return true;
      }
    }
    this.log(`Authentication failed. Status: ${status}, Response: ${JSON.stringify(data)}`);
    return false;
  }

  // Helper parsers
  public parseTimeFromExcel(timeValue: any): { hours: number; minutes: number; seconds: number } {
    if (!timeValue) {
      return { hours: 0, minutes: 0, seconds: 0 };
    }

    // Handle string formatting
    if (typeof timeValue === 'string') {
      const cleanTime = timeValue.trim().toUpperCase();
      // Try formats
      const regexes = [
        /^(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)?$/,
        /^(\d{1,2}):(\d{2})\s*(AM|PM)?$/,
      ];

      for (const regex of regexes) {
        const match = cleanTime.match(regex);
        if (match) {
          let hours = parseInt(match[1], 10);
          const minutes = parseInt(match[2], 10);
          const seconds = match[3] ? parseInt(match[3], 10) : 0;
          const ampm = match[match.length - 1];

          if (ampm === 'PM' && hours < 12) hours += 12;
          if (ampm === 'AM' && hours === 12) hours = 0;

          return { hours, minutes, seconds };
        }
      }
      return { hours: 0, minutes: 0, seconds: 0 };
    }

    // Handle JS Date object
    if (timeValue instanceof Date) {
      return {
        hours: timeValue.getHours(),
        minutes: timeValue.getMinutes(),
        seconds: timeValue.getSeconds(),
      };
    }

    // Handle fractional day number (e.g. 0.5 is 12:00:00)
    if (typeof timeValue === 'number') {
      if (timeValue >= 0 && timeValue <= 1) {
        const totalSeconds = Math.floor(timeValue * 86400);
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;
        return { hours, minutes, seconds };
      }
    }

    return { hours: 0, minutes: 0, seconds: 0 };
  }

  public parseEstimatedTime(timeValue: any): number {
    if (!timeValue) {
      return 3600; // default 1 hour
    }

    const timeStr = String(timeValue).trim().toLowerCase();
    const match = timeStr.match(/(\d+(?:\.\d+)?)\s*(hr|hrs|hour|hours|min|mins|minutes?)/);

    if (match) {
      const value = parseFloat(match[1]);
      const unit = match[2];
      if (unit.startsWith('hr') || unit.startsWith('hour')) {
        return Math.floor(value * 3600);
      } else {
        return Math.floor(value * 60);
      }
    }

    const numberMatch = timeStr.match(/(\d+(?:\.\d+)?)/);
    if (numberMatch) {
      return Math.floor(parseFloat(numberMatch[1]) * 3600);
    }

    return 3600;
  }

  // Preloaders
  public async preloadGeneralCache() {
    this.log('Preloading general cache (clients, work order types, tags)...');
    
    // Clients
    let res = await this.requestJson('GET', 'clients/');
    if (res.status === 200) {
      const items = res.data?.data?.items || [];
      items.forEach((item: any) => {
        this.cache.clients[item.name] = item.id;
      });
      this.log(`Loaded ${items.length} clients into cache.`);
    }

    // Work Order Types
    res = await this.requestJson('GET', 'work-order-types/');
    if (res.status === 200) {
      const items = res.data?.data?.items || [];
      items.forEach((item: any) => {
        this.cache.work_order_types[item.name.trim()] = item.id;
      });
      this.log(`Loaded ${items.length} work order types into cache.`);
    }

    // Tags
    res = await this.requestJson('GET', 'tags/');
    if (res.status === 200) {
      const items = res.data?.data?.items || [];
      items.forEach((item: any) => {
        this.cache.tags[item.name.trim()] = item.id;
      });
      this.log(`Loaded ${items.length} tags into cache.`);
    }
  }

  public async preloadCacheForClient(clientId: string, clientName: string) {
    this.log(`Preloading cache for client: ${clientName}...`);
    const endpoints = [
      { key: 'forms', path: 'forms' },
      { key: 'assets', path: 'assets' },
      { key: 'teams', path: 'teams' },
      { key: 'locations', path: 'locations' },
    ] as const;

    for (const ep of endpoints) {
      const res = await this.requestJson('GET', `${ep.path}/`, null, { client: clientId });
      if (res.status === 200) {
        const items = res.data?.data?.items || [];
        items.forEach((item: any) => {
          const lookupKey = `${item.name}_${clientId}`;
          this.cache[ep.key][lookupKey] = item.id;
        });
      }
    }
    this.log(`Client cache loaded for ${clientName}.`);
  }

  // Cached lookup
  private async getCachedOrFetch(
    cacheKey: keyof typeof this.cache,
    searchValue: string,
    endpoint: string,
    clientId?: string
  ): Promise<string | null> {
    const lookupKey = clientId ? `${searchValue}_${clientId}` : searchValue;

    if (this.cache[cacheKey][lookupKey]) {
      return this.cache[cacheKey][lookupKey];
    }

    const params: Record<string, any> = { search: searchValue };
    if (clientId) {
      params.client = clientId;
    }

    const res = await this.requestJson('GET', `${endpoint}/`, null, params);
    if (res.status === 200) {
      let items = res.data?.data?.items || [];

      // Filter by client if required
      if (items.length > 0 && clientId) {
        items = items.filter((item: any) => {
          const clientRef = item.client || item.client_id;
          if (typeof clientRef === 'object' && clientRef !== null) {
            return clientRef.id === clientId;
          }
          return clientRef === clientId;
        });
      }

      if (items.length > 0) {
        // Find exact match
        const exactMatch = items.find(
          (item: any) => item.name?.toLowerCase() === searchValue.toLowerCase()
        );
        const selected = exactMatch || items[0];
        this.cache[cacheKey][lookupKey] = selected.id;
        return selected.id;
      }
    }

    // Location search fallbacks (as in the script)
    if (cacheKey === 'locations' && searchValue) {
      if (searchValue.includes('(')) {
        const mainPart = searchValue.split('(')[0].trim();
        const fallbackRes = await this.requestJson('GET', `${endpoint}/`, null, {
          search: mainPart,
          client: clientId,
        });
        if (fallbackRes.status === 200) {
          const items = fallbackRes.data?.data?.items || [];
          if (items.length > 0) {
            this.cache.locations[lookupKey] = items[0].id;
            return items[0].id;
          }
        }
      }

      // Strip numbers/periods/symbols
      const cleanName = searchValue
        .replace(/[0-9.]+/g, '')
        .replace(/[-_()]+/g, ' ')
        .trim();

      if (cleanName !== searchValue) {
        const fallbackRes = await this.requestJson('GET', `${endpoint}/`, null, {
          search: cleanName,
          client: clientId,
        });
        if (fallbackRes.status === 200) {
          const items = fallbackRes.data?.data?.items || [];
          if (items.length > 0) {
            this.cache.locations[lookupKey] = items[0].id;
            return items[0].id;
          }
        }
      }
    }

    return null;
  }

  // Create Work Order
  public async createWorkOrder(row: Record<string, any>): Promise<boolean> {
    const clientName = String(row['Client'] || '').trim();
    if (!clientName) {
      this.log('Skipping row: Client name is missing.');
      return false;
    }

    const clientId = await this.getCachedOrFetch('clients', clientName, 'clients');
    if (!clientId) {
      this.log(`Failed to find client: ${clientName}`);
      return false;
    }

    let locationCode = String(row['Code'] || '').trim();
    const locationName = String(row['Location name'] || '').trim();

    if (locationCode.endsWith('.0')) {
      locationCode = locationCode.slice(0, -2);
    }

    const searchParts = [locationName, locationCode].filter(
      (x) => x && x !== '0' && x !== 'nan' && x !== 'NaN'
    );
    if (searchParts.length === 0) {
      this.log(`No valid location search parts found for row client: ${clientName}`);
      return false;
    }

    const searchStr = searchParts.join(' ');
    const locationId = await this.getCachedOrFetch('locations', searchStr, 'locations', clientId);
    if (!locationId) {
      this.log(`Failed to find location: "${searchStr}" for client: ${clientName}`);
      return false;
    }

    const checklistFormName = String(row['Checklist Form'] || '').trim();
    const checklistFormId = await this.getCachedOrFetch('forms', checklistFormName, 'forms', clientId);
    if (!checklistFormId) {
      this.log(`Failed to find checklist form: "${checklistFormName}" for client: ${clientName}`);
      return false;
    }

    const workOrderTypeName = String(row['Work Order Type'] || '').trim();
    const workOrderTypeId = await this.getCachedOrFetch('work_order_types', workOrderTypeName, 'work-order-types');
    if (!workOrderTypeId) {
      this.log(`Failed to find work order type: "${workOrderTypeName}" (searched globally)`);
      return false;
    }

    const teamName = String(row['Team Name'] || '').trim();
    let teamId: string | null = null;
    if (teamName && teamName !== '0' && teamName !== 'nan' && teamName !== 'NaN') {
      teamId = await this.getCachedOrFetch('teams', teamName, 'teams', clientId);
      if (!teamId) {
        this.log(`Warning: Failed to find team "${teamName}" for client: ${clientName} - Proceeding without team`);
      }
    }

    // Parse Dates/Times
    let dueVal = row['Due Date'];
    let timeVal = row['Due Time'];

    let dueDateObj: Date;
    if (dueVal instanceof Date) {
      dueDateObj = dueVal;
    } else if (typeof dueVal === 'number') {
      // Excel serial date format
      dueDateObj = new Date(Math.round((dueVal - 25569) * 86400 * 1000));
    } else if (typeof dueVal === 'string') {
      const cleanVal = dueVal.trim();
      if (cleanVal.includes('-')) {
        dueDateObj = new Date(cleanVal);
      } else if (cleanVal.includes('/')) {
        const parts = cleanVal.split('/');
        // Format d/m/yyyy -> parts[2] is year, parts[1] is month (0-indexed), parts[0] is day
        dueDateObj = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      } else {
        this.log(`Invalid due date format: ${dueVal}`);
        return false;
      }
    } else {
      this.log(`Invalid due date format: ${dueVal}`);
      return false;
    }

    const timeObj = this.parseTimeFromExcel(timeVal);
    dueDateObj.setHours(timeObj.hours);
    dueDateObj.setMinutes(timeObj.minutes);
    dueDateObj.setSeconds(timeObj.seconds);

    // Convert Asia/Kolkata (+5:30) to UTC (Z)
    // We add timezone adjustment manually or parse it
    // Let's create an offset conversion: Date object represents local time in Asia/Kolkata
    // We construct the string and treat it as a localized date, then convert to UTC ISO string.
    const year = dueDateObj.getFullYear();
    const month = String(dueDateObj.getMonth() + 1).padStart(2, '0');
    const day = String(dueDateObj.getDate()).padStart(2, '0');
    const hours = String(dueDateObj.getHours()).padStart(2, '0');
    const minutes = String(dueDateObj.getMinutes()).padStart(2, '0');
    const seconds = String(dueDateObj.getSeconds()).padStart(2, '0');

    // Create an ISO string with +05:30 offset
    const kolkataIsoStr = `${year}-${month}-${day}T${hours}:${minutes}:${seconds}+05:30`;
    const finalDate = new Date(kolkataIsoStr);
    
    if (isNaN(finalDate.getTime())) {
      this.log(`Error parsing due date/time combo into date.`);
      return false;
    }

    const formattedDueAt = finalDate.toISOString(); // This is in UTC 'Z'

    const estTimeSeconds = this.parseEstimatedTime(row['Estimated Time']);

    const assetName = String(row['Asset'] || '').trim();
    let assetId: string | null = null;
    if (assetName && assetName !== '0' && assetName !== 'nan' && assetName !== 'NaN') {
      let cleanAsset = assetName;
      if (assetName.endsWith('.0')) {
        cleanAsset = assetName.slice(0, -2);
      }
      assetId = await this.getCachedOrFetch('assets', cleanAsset, 'assets', clientId);
    }

    const tagIds: string[] = [];
    const tagsFromExcel = String(row['Tags'] || '').trim();
    if (tagsFromExcel && tagsFromExcel !== '0' && tagsFromExcel !== 'nan' && tagsFromExcel !== 'NaN') {
      const tagNames = tagsFromExcel.split(',').map((t) => t.trim()).filter(Boolean);
      for (const name of tagNames) {
        const id = await this.getCachedOrFetch('tags', name, 'tags');
        if (id) {
          tagIds.push(String(id));
        }
      }
    }

    const schedule = String(row['Schedule'] || '').trim().toLowerCase();
    const isRecurring = schedule === 'daily' || schedule === 'weekly';

    const checklistReq = String(row['Checklist required for completion'] || '').trim().toUpperCase();
    const isChecklistRequired = ['Y', 'YES', '1', 'TRUE'].includes(checklistReq);

    const payload: Record<string, any> = {
      client: clientId,
      location: locationId,
      title: row['Title'] || 'Work Order',
      teams: teamId ? [teamId] : [],
      work_order_type: workOrderTypeId,
      description: '',
      status: 'open',
      priority: 'high',
      due_at: formattedDueAt,
      estimated_time: String(estTimeSeconds),
      auto_create_chain: isRecurring,
      checklist_form: checklistFormId,
      assignees: [],
      is_checklist_form_required: isChecklistRequired,
    };

    if (assetId) {
      payload.asset = assetId;
    }

    if (tagIds.length > 0) {
      payload.tags = tagIds;
    }

    if (isRecurring) {
      if (schedule === 'daily') {
        payload.recurrence = { interval: 1, precision: [1, 2, 3, 4, 5] };
        payload.recurrence_type = 'daily';
      } else if (schedule === 'weekly') {
        payload.recurrence = { interval: 1, precision: [1, 2, 3, 4, 5] };
        payload.recurrence_type = 'weekly';
      }
    }

    const res = await this.requestJson('POST', 'work-orders/', payload);
    if (res.status === 201) {
      this.log(`Created Work Order successfully: "${row['Title']}"`);
      return true;
    } else {
      this.log(`Failed to create work order: "${row['Title']}". Status: ${res.status}, Response: ${JSON.stringify(res.data)}`);
      return false;
    }
  }

  // Delete helpers
  public async getClientIdByName(clientName: string): Promise<string | null> {
    if (!clientName) {
      return null;
    }
    const res = await this.requestJson('GET', 'clients/', null, { search: clientName });
    if (res.status === 200) {
      const items = res.data?.data?.items || [];
      const exact = items.find((i: any) => i.name?.trim().toLowerCase() === clientName.trim().toLowerCase());
      const selected = exact || items[0];
      if (selected) {
        this.log(`Selected client: ${selected.name} (${selected.id})`);
        return selected.id;
      }
    }
    return null;
  }

  public async fetchLocationsForClient(clientId: string): Promise<SiteLocation[]> {
    const locations: SiteLocation[] = [];
    const seen = new Set<string>();
    const pageSize = 100;

    for (let page = 1; page <= 100; page++) {
      const res = await this.requestJson('GET', 'locations/', null, {
        client: clientId,
        page,
        page_size: pageSize,
      });

      if (res.status !== 200) {
        this.log(`Failed to fetch locations. Status: ${res.status}`);
        return locations;
      }

      const items = res.data?.data?.items || [];
      for (const item of items) {
        if (item.id && !seen.has(item.id)) {
          seen.add(item.id);
          locations.push({
            id: item.id,
            name: item.name || item.code || item.id,
            code: item.code,
          });
        }
      }

      if (items.length < pageSize) {
        break;
      }
    }

    this.log(`Loaded ${locations.length} site locations for client ${clientId}.`);
    return locations;
  }

  public async fetchWorkOrders(clientId: string, limit: number, onlyOverdue: boolean, locationId?: string): Promise<any[]> {
    const workOrders: any[] = [];
    const pageSize = Math.min(100, Math.max(1, limit));

    for (let page = 1; workOrders.length < limit; page++) {
      const params: Record<string, any> = {
        client: clientId,
        page,
        page_size: Math.min(pageSize, limit - workOrders.length),
      };
      if (onlyOverdue) {
        params.is_overdue = 'true';
      }
      if (locationId) {
        params.location = locationId;
      }

      const res = await this.requestJson('GET', 'work-orders/', null, params);
      if (res.status !== 200) {
        this.log(`Failed to fetch work orders. Status: ${res.status}`);
        return workOrders;
      }

      const items = res.data?.data?.items || [];
      workOrders.push(...items);

      if (items.length < params.page_size) {
        break;
      }
    }

    return workOrders;
  }

  public async findWorkOrderByNumber(clientId: string, workOrderNumber: string): Promise<any | null> {
    const expected = workOrderNumber.trim().toLowerCase();
    const res = await this.requestJson('GET', 'work-orders/', null, {
      client: clientId,
      search: workOrderNumber,
      page: 1,
      page_size: 10,
    });

    if (res.status === 200) {
      const items = res.data?.data?.items || [];
      for (const item of items) {
        const itemNum = String(item.work_order_number || '').trim().toLowerCase();
        if (itemNum === expected) {
          return item;
        }
      }
    }
    return null;
  }

  public async resolveLocationId(clientId: string, locationValue: string): Promise<string | null> {
    if (!locationValue) return '';
    const res = await this.requestJson('GET', 'locations/', null, {
      client: clientId,
      search: locationValue,
      page: 1,
      page_size: 10,
    });

    if (res.status === 200) {
      const items = res.data?.data?.items || [];
      const expected = locationValue.trim().toLowerCase();
      
      for (const item of items) {
        const candidates = [item.id, item.name, item.code];
        if (candidates.some((c) => String(c || '').trim().toLowerCase() === expected)) {
          return item.id;
        }
      }

      if (items.length === 1) {
        return items[0].id;
      }
    }
    return null;
  }

  public async locationMatches(clientId: string, workOrder: any, expectedLocation: string): Promise<boolean> {
    if (!expectedLocation) return true;
    const resolvedId = await this.resolveLocationId(clientId, expectedLocation);
    if (!resolvedId) return false;

    const woLocation = workOrder.location;
    const woLocationId = typeof woLocation === 'object' && woLocation !== null ? woLocation.id : woLocation;
    return String(woLocationId || '').toLowerCase() === String(resolvedId).toLowerCase();
  }

  public async deleteWorkOrder(workOrder: any, dryRun: boolean): Promise<boolean> {
    const id = workOrder.id;
    const title = workOrder.title || '';
    const number = workOrder.work_order_number || '';

    if (dryRun) {
      this.log(`DRY RUN - would delete: ${id} | ${number} | ${title}`);
      return true;
    }

    const res = await this.requestJson('DELETE', `work-orders/${id}/`);
    if (res.status === 200) {
      this.log(`Deleted: ${id} | ${number} | ${title}`);
      return true;
    }

    this.log(`Failed delete: ${id}. Status: ${res.status}, Response: ${JSON.stringify(res.data)}`);
    return false;
  }

  // ================= TICKET BULK UPLOAD HELPERS =================

  public async getTicketLocationId(searchValue: string, clientId: string): Promise<string | null> {
    if (!searchValue) return null;

    const cleanSearchValue = String(searchValue).trim();
    const cacheKey = `${cleanSearchValue}_${clientId}`;

    if (this.cache.locations[cacheKey]) {
      return this.cache.locations[cacheKey];
    }

    const params = {
      search: cleanSearchValue,
      client: clientId
    };

    const res = await this.requestJson('GET', 'locations/', null, params);
    if (res.status === 200) {
      const items = res.data?.data?.items || [];
      if (items.length > 0) {
        const locId = items[0].id;
        this.cache.locations[cacheKey] = locId;
        this.log(`Found location: ${cleanSearchValue}`);
        return locId;
      }
    } else {
      this.log(`❌ Location API failed: ${res.status}`);
      return null;
    }

    // Fallback 1: Replace underscores with spaces
    if (cleanSearchValue.includes('_')) {
      const spaceVal = cleanSearchValue.replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
      this.log(`🔁 Retry location (replace underscores): "${spaceVal}"`);
      const fallbackId = await this.getTicketLocationId(spaceVal, clientId);
      if (fallbackId) {
        this.cache.locations[cacheKey] = fallbackId;
        return fallbackId;
      }
    }

    // Fallback 2: Remove brackets
    if (cleanSearchValue.includes('(')) {
      const shortVal = cleanSearchValue.split('(')[0].trim();
      this.log(`🔁 Retry location (remove brackets): "${shortVal}"`);
      const fallbackId = await this.getTicketLocationId(shortVal, clientId);
      if (fallbackId) {
        this.cache.locations[cacheKey] = fallbackId;
        return fallbackId;
      }
    }

    // Fallback 3: Clean numbers/symbols
    const cleanName = cleanSearchValue
      .replace(/[0-9.]+/g, '')
      .replace(/[-_()]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    if (cleanName !== cleanSearchValue && cleanName.length > 0) {
      this.log(`🔁 Retry location (cleaned): "${cleanName}"`);
      const fallbackId = await this.getTicketLocationId(cleanName, clientId);
      if (fallbackId) {
        this.cache.locations[cacheKey] = fallbackId;
        return fallbackId;
      }
    }

    // Fallback 4: Short search
    const words = cleanSearchValue.split(/\s+/);
    if (words.length > 2) {
      const shortVal = words.slice(0, 2).join(' ');
      this.log(`🔁 Retry location (short): "${shortVal}"`);
      const fallbackId = await this.getTicketLocationId(shortVal, clientId);
      if (fallbackId) {
        this.cache.locations[cacheKey] = fallbackId;
        return fallbackId;
      }
    }

    this.log(`❌ Location not found: "${cleanSearchValue}"`);
    return null;
  }

  public async getTicketMemberId(searchValue: string): Promise<string | null> {
    if (!searchValue) return null;

    const cleanSearchValue = String(searchValue).trim().toLowerCase();

    if (this.cache.users[cleanSearchValue]) {
      return this.cache.users[cleanSearchValue];
    }

    const res = await this.requestJson('GET', 'members/', null, { search: cleanSearchValue });
    if (res.status !== 200) {
      this.log(`❌ Members API failed: ${res.status}`);
      return null;
    }

    const items = res.data?.data?.items || [];
    for (const item of items) {
      const user = item.user || {};
      const firstName = user.first_name || '';
      const lastName = user.last_name || '';
      const fullName = `${firstName} ${lastName}`.trim().toLowerCase();

      const searchWords = cleanSearchValue.split(/\s+/);
      const allWordsMatched = searchWords.every(word => fullName.includes(word));

      if (allWordsMatched) {
        const userId = user.id;
        this.cache.users[cleanSearchValue] = userId;
        this.log(`Found owner: ${firstName} ${lastName}`);
        return userId;
      }
    }

    this.log(`❌ Owner not found: ${cleanSearchValue}`);
    return null;
  }

  public async createTicket(
    row: Record<string, any>,
    clientId: string,
    ticketTypeId: string,
    priorityId: string,
    dryRun: boolean
  ): Promise<boolean> {
    const normalisedRow: Record<string, any> = {};
    Object.keys(row).forEach(key => {
      normalisedRow[key.trim().toLowerCase()] = row[key];
    });

    let subject = String(normalisedRow['subject'] || '').trim();
    if (!subject || subject.toLowerCase() === 'nan') {
      subject = 'Auto Ticket';
    }

    const locationVal = normalisedRow['location'];
    const ownerVal = normalisedRow['owner'];

    const locationId = await this.getTicketLocationId(locationVal, clientId);
    const ownerId = await this.getTicketMemberId(ownerVal);

    if (!locationId) {
      this.log(`❌ Location not found: "${locationVal}"`);
      return false;
    }

    if (!ownerId) {
      this.log(`❌ Owner not found: "${ownerVal}"`);
      return false;
    }

    const tagsVal = normalisedRow['tags'];
    const tagIds: number[] = [];
    if (tagsVal && String(tagsVal).trim() !== '' && String(tagsVal).toLowerCase() !== 'nan') {
      const tagList = String(tagsVal).split(',').map(t => t.trim()).filter(Boolean);
      for (const name of tagList) {
        const id = await this.getTicketTagId(name);
        if (id !== null) {
          tagIds.push(id);
        }
      }
    }

    const payload = {
      client: clientId,
      location: locationId,
      type: ticketTypeId,
      priority: priorityId,
      status: 'open',
      category: 'general',
      subject: subject,
      description: subject,
      tags: tagIds,
      l1_assignee: ownerId,
    };

    if (dryRun) {
      this.log(`[DRY RUN] Would create Ticket: "${subject}" | Location: "${locationVal}" (ID: ${locationId}) | Owner: "${ownerVal}" (ID: ${ownerId}) | Tags: ${JSON.stringify(tagIds)}`);
      return true;
    }

    const res = await this.requestJson('POST', 'tickets/', payload);
    if (res.status === 200 || res.status === 201) {
      this.log(`✅ Created → Subject: "${subject}" | Location: "${locationVal}" (ID: ${locationId}) | Owner: "${ownerVal}" (ID: ${ownerId}) | Tags: ${JSON.stringify(tagIds)}`);
      return true;
    } else {
      const responseText = typeof res.data === 'object' ? JSON.stringify(res.data) : String(res.data || '');
      this.log(`❌ Failed → Subject: "${subject}" | Location: "${locationVal}" (ID: ${locationId}) | Owner: "${ownerVal}" (ID: ${ownerId}) | Tags: ${JSON.stringify(tagIds)}. Status: ${res.status}, Response: ${responseText}`);
      return false;
    }
  }

  public async getTicketTagId(tagName: string): Promise<number | null> {
    if (!tagName) return null;
    const cleanTagName = tagName.trim();
    
    // Check cache
    const cacheKey = `ticket_tag_${cleanTagName.toLowerCase()}`;
    if (this.cache.tags[cacheKey]) {
      return Number(this.cache.tags[cacheKey]);
    }

    // If already an integer ID
    if (/^\d+$/.test(cleanTagName)) {
      return parseInt(cleanTagName, 10);
    }

    const res = await this.requestJson('GET', 'tags/', null, { search: cleanTagName });
    if (res.status === 200) {
      const items = res.data?.data?.items || res.data?.items || [];
      const exactMatch = items.find((item: any) => String(item.name).toLowerCase() === cleanTagName.toLowerCase());
      const selected = exactMatch || items[0];
      if (selected && selected.id !== undefined) {
        this.cache.tags[cacheKey] = String(selected.id);
        this.log(`Resolved tag: "${cleanTagName}" to ID ${selected.id}`);
        return Number(selected.id);
      }
    }

    this.log(`⚠️ Warning: Tag "${cleanTagName}" not found in API.`);
    return null;
  }

  public async fetchConfigDetails(clientId: string, ticketTypeId: string, priorityId: string) {
    let clientName = 'Unknown';
    let ticketTypeName = 'Unknown';
    let priorityName = 'Unknown';

    if (clientId) {
      const res = await this.requestJson('GET', `clients/${clientId}/`);
      if (res.status === 200) {
        clientName = res.data?.data?.name || res.data?.name || 'Unknown';
      }
    }

    if (ticketTypeId) {
      const res = await this.requestJson('GET', `ticket-types/${ticketTypeId}/`);
      if (res.status === 200) {
        ticketTypeName = res.data?.data?.name || res.data?.name || 'Unknown';
      }
    }

    if (priorityId) {
      const res = await this.requestJson('GET', `ticket-priorities/${priorityId}/`);
      if (res.status === 200) {
        priorityName = res.data?.data?.name || res.data?.name || 'Unknown';
      }
    }

    return { clientName, ticketTypeName, priorityName };
  }
}
