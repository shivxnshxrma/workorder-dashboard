'use client';

import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload,
  Trash2,
  Terminal,
  FileText,
  LogOut,
  User,
  Search,
  Database,
  CheckCircle,
  AlertTriangle,
  Play,
  X,
  FileSpreadsheet,
  Settings,
  Moon,
  Info,
  Clock,
  ShieldAlert,
  MapPin
} from 'lucide-react';
import { MerlinClient, type SiteLocation } from '@/lib/merlin';

interface AuditLog {
  id: string;
  timestamp: string;
  user: string;
  action: 'upload' | 'delete';
  source: string;
  total: number;
  success: number;
  failed: number;
}

interface LogLine {
  type: 'success' | 'warning' | 'danger' | 'info' | 'debug' | 'general';
  text: string;
}

const MAX_DELETE_WORKERS = 10;

export default function DashboardClient({ userEmail, onLogout }: { userEmail: string; onLogout: () => void }) {
  const [activeTab, setActiveTab] = useState<'upload' | 'ticket_upload' | 'delete' | 'logs'>('upload');
  
  // App Configs
  const [uploadApiUrl, setUploadApiUrl] = useState('https://api-mneo-cbre.merlin-soteria.in/api/v1/');
  const [uploadUsername, setUploadUsername] = useState('');
  const [uploadPassword, setUploadPassword] = useState('');
  
  const [deleteApiUrl, setDeleteApiUrl] = useState('https://api-merlin.tenonfm-india.com/api/v1/');
  const [deleteUsername, setDeleteUsername] = useState('');
  const [deletePassword, setDeletePassword] = useState('');

  // Ticket Upload Configs & State (using credentials from ticket_bulk_upload.txt)
  const [ticketApiUrl, setTicketApiUrl] = useState('https://api-mneo.soteria.in/api/v1/');
  const [ticketEmail, setTicketEmail] = useState('anuj.mourya@soteria.in');
  const [ticketPassword, setTicketPassword] = useState('Anuj2312@');
  const [ticketClientId, setTicketClientId] = useState('a699bab9-3d3a-4ca4-b572-f6383cf47233');
  const [ticketTypeId, setTicketTypeId] = useState('7150d7cd-845a-4900-a316-21bbf6ab662e');
  const [ticketPriorityId, setTicketPriorityId] = useState('5eb3ecec-9303-4900-8a83-a081f8f1ef25');
  const [ticketFile, setTicketFile] = useState<File | null>(null);
  const [ticketRows, setTicketRows] = useState<any[]>([]);
  
  // Upload State
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [uploadRows, setUploadRows] = useState<any[]>([]);
  
  // Delete State
  const [deleteFile, setDeleteFile] = useState<File | null>(null);
  const [deleteRows, setDeleteRows] = useState<any[]>([]);
  const [deleteMode, setDeleteMode] = useState<'sheet' | 'client_all'>('sheet');
  const [deleteLimit, setDeleteLimit] = useState(10);
  const [dryRun, setDryRun] = useState(true);
  const [deleteOnlyOverdue, setDeleteOnlyOverdue] = useState(false);
  const [verifyLocation, setVerifyLocation] = useState(true);
  const [deleteClientName, setDeleteClientName] = useState('');
  const [deleteClientId, setDeleteClientId] = useState('');
  const [deleteLocations, setDeleteLocations] = useState<SiteLocation[]>([]);
  const [selectedDeleteLocationId, setSelectedDeleteLocationId] = useState('');
  const [locationsClientId, setLocationsClientId] = useState('');
  const [isLoadingLocations, setIsLoadingLocations] = useState(false);
  
  // Process State
  const [isRunning, setIsRunning] = useState(false);
  const [isStopping, setIsStopping] = useState(false);
  const [progress, setProgress] = useState(0);
  const [stats, setStats] = useState({ total: 0, current: 0, success: 0, failed: 0 });
  const [logs, setLogs] = useState<LogLine[]>([]);
  
  // Audit Logs State
  const [auditLogs, setAuditLogs] = useState<AuditLog[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  
  const consoleBoxRef = useRef<HTMLDivElement>(null);
  const stopRequestedRef = useRef(false);

  // Load audit logs from localStorage on mount
  useEffect(() => {
    const saved = localStorage.getItem('soteria_audit_logs');
    if (saved) {
      try {
        setAuditLogs(JSON.parse(saved));
      } catch (e) {
        console.error('Failed to parse audit logs', e);
      }
    }
  }, []);

  // Keep console stable by scrolling the box to the bottom without scrolling the main page window.
  useEffect(() => {
    if (consoleBoxRef.current) {
      consoleBoxRef.current.scrollTop = consoleBoxRef.current.scrollHeight;
    }
  }, [logs]);

  const addLog = (text: string, type: LogLine['type'] = 'general') => {
    const timestamp = new Date().toLocaleTimeString();
    setLogs((prev) => [...prev, { type, text: `[${timestamp}] ${text}` }]);
  };

  const clearConsole = () => {
    setLogs([]);
  };

  const requestStop = () => {
    stopRequestedRef.current = true;
    setIsStopping(true);
    addLog('Stop requested. Current in-flight request(s) will finish; no new work will start.', 'warning');
  };

  const resetDeleteLocations = () => {
    setDeleteLocations([]);
    setSelectedDeleteLocationId('');
    setLocationsClientId('');
  };

  // Handle excel/csv uploads for creation
  const handleUploadFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploadFile(file);
    
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet);
        setUploadRows(rows);
        addLog(`File loaded: "${file.name}" with ${rows.length} rows.`, 'info');
      } catch (err: any) {
        addLog(`Failed to parse file: ${err.message}`, 'danger');
      }
    };
    reader.readAsBinaryString(file);
  };

  // Handle excel/csv uploads for ticket creation
  const handleTicketFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setTicketFile(file);
    
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet);
        setTicketRows(rows);
        addLog(`Ticket file loaded: "${file.name}" with ${rows.length} rows.`, 'info');
      } catch (err: any) {
        addLog(`Failed to parse file: ${err.message}`, 'danger');
      }
    };
    reader.readAsBinaryString(file);
  };

  // Handle excel/csv uploads for deletion
  const handleDeleteFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setDeleteFile(file);
    
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet);
        setDeleteRows(rows);
        addLog(`Delete file loaded: "${file.name}" with ${rows.length} rows.`, 'info');
      } catch (err: any) {
        addLog(`Failed to parse file: ${err.message}`, 'danger');
      }
    };
    reader.readAsBinaryString(file);
  };

  const saveAuditLog = (action: 'upload' | 'delete', source: string, total: number, success: number, failed: number) => {
    const newLog: AuditLog = {
      id: Math.random().toString(36).substring(2, 9),
      timestamp: new Date().toISOString(),
      user: userEmail,
      action,
      source,
      total,
      success,
      failed
    };
    const updated = [newLog, ...auditLogs];
    setAuditLogs(updated);
    localStorage.setItem('soteria_audit_logs', JSON.stringify(updated));
  };

  const loadDeleteLocations = async () => {
    const clientId = deleteClientId.trim();
    if (!clientId) {
      addLog('Client ID is required to load site locations.', 'warning');
      return;
    }

    setIsLoadingLocations(true);
    resetDeleteLocations();
    addLog(`Loading site locations for client: ${clientId}...`, 'info');

    const client = new MerlinClient(
      {
        apiUrl: deleteApiUrl,
        username: deleteUsername,
        password: deletePassword,
      },
      (msg) => {
        if (msg.includes('failed') || msg.includes('Failed')) {
          addLog(msg, 'danger');
        } else {
          addLog(msg, 'general');
        }
      }
    );

    const authenticated = await client.authenticate();
    if (!authenticated) {
      addLog('Authentication failed. Could not load locations.', 'danger');
      setIsLoadingLocations(false);
      return;
    }

    const locations = await client.fetchLocationsForClient(clientId);
    setDeleteLocations(locations);
    setLocationsClientId(clientId);
    if (locations.length === 0) {
      addLog('No site locations found for this Client ID.', 'warning');
    } else {
      addLog(`Location dropdown ready with ${locations.length} site locations.`, 'success');
    }
    setIsLoadingLocations(false);
  };

  const startBulkUpload = async () => {
    if (uploadRows.length === 0) {
      addLog('No rows to upload. Please upload an Excel/CSV file first.', 'warning');
      return;
    }

    stopRequestedRef.current = false;
    setIsStopping(false);
    setIsRunning(true);
    setProgress(0);
    setStats({ total: uploadRows.length, current: 0, success: 0, failed: 0 });
    clearConsole();

    addLog('Starting bulk upload sequence...', 'info');

    const client = new MerlinClient(
      {
        apiUrl: uploadApiUrl,
        username: uploadUsername,
        password: uploadPassword,
      },
      (msg) => {
        // Parse types based on content to make logs beautiful
        if (msg.includes('success') || msg.includes('successful')) {
          addLog(msg, 'success');
        } else if (msg.includes('Warning') || msg.includes('warning') || msg.includes('Skipping')) {
          addLog(msg, 'warning');
        } else if (msg.includes('Failed') || msg.includes('failed') || msg.includes('Error')) {
          addLog(msg, 'danger');
        } else if (msg.includes('Preloading') || msg.includes('Preloaded') || msg.includes('Loaded')) {
          addLog(msg, 'info');
        } else if (msg.includes('DEBUG')) {
          addLog(msg, 'debug');
        } else {
          addLog(msg, 'general');
        }
      }
    );

    const authenticated = await client.authenticate();
    if (!authenticated) {
      addLog('Authentication failed. Terminating upload.', 'danger');
      setIsRunning(false);
      setIsStopping(false);
      return;
    }

    if (stopRequestedRef.current) {
      addLog('Upload stopped before cache preload.', 'warning');
      setIsRunning(false);
      setIsStopping(false);
      return;
    }

    await client.preloadGeneralCache();

    // Group client caching
    const uniqueClients = Array.from(new Set(uploadRows.map((r: any) => String(r.Client || '').trim()).filter(Boolean)));
    for (const clientName of uniqueClients) {
      if (stopRequestedRef.current) {
        addLog('Upload stopped during client cache preload.', 'warning');
        break;
      }
      const clientId = await client.getClientIdByName(clientName);
      if (clientId) {
        await client.preloadCacheForClient(clientId, clientName);
      }
    }

    let successCount = 0;
    let failedCount = 0;

    for (let i = 0; i < uploadRows.length; i++) {
      if (stopRequestedRef.current) {
        addLog(`Upload stopped. ${uploadRows.length - i} row(s) left unprocessed.`, 'warning');
        break;
      }

      const row = uploadRows[i];
      addLog(`Processing row ${i + 1}/${uploadRows.length}: "${row.Title || 'Untitled Work Order'}"...`, 'general');
      
      try {
        const success = await client.createWorkOrder(row);
        if (success) {
          successCount++;
        } else {
          failedCount++;
        }
      } catch (err: any) {
        addLog(`Exception creating work order in row ${i + 1}: ${err.message}`, 'danger');
        failedCount++;
      }

      const current = i + 1;
      setStats({
        total: uploadRows.length,
        current,
        success: successCount,
        failed: failedCount,
      });
      setProgress(Math.round((current / uploadRows.length) * 100));
    }

    const uploadStopped = stopRequestedRef.current;
    addLog(`\n--- Upload Process ${uploadStopped ? 'Stopped' : 'Completed'} ---`, uploadStopped ? 'warning' : 'info');
    addLog(`Total Processed: ${uploadRows.length}`, 'general');
    addLog(`Successfully Created: ${successCount}`, 'success');
    addLog(`Failed: ${failedCount}`, 'danger');

    saveAuditLog('upload', uploadFile?.name || 'Excel/CSV file', uploadRows.length, successCount, failedCount);
    stopRequestedRef.current = false;
    setIsRunning(false);
    setIsStopping(false);
  };

  const startTicketBulkUpload = async (isDryRun: boolean) => {
    if (ticketRows.length === 0) {
      addLog('No rows to upload. Please upload an Excel/CSV file first.', 'warning');
      return;
    }

    stopRequestedRef.current = false;
    setIsStopping(false);
    setIsRunning(true);
    setProgress(0);
    setStats({ total: ticketRows.length, current: 0, success: 0, failed: 0 });
    clearConsole();

    addLog(`Starting ticket bulk upload sequence ${isDryRun ? '(DRY RUN)' : ''}...`, 'info');

    const client = new MerlinClient(
      {
        apiUrl: ticketApiUrl,
        username: ticketEmail,
        password: ticketPassword,
      },
      (msg) => {
        if (msg.includes('success') || msg.includes('successful') || msg.includes('Created →') || msg.includes('Found owner') || msg.includes('Found location')) {
          addLog(msg, 'success');
        } else if (msg.includes('Warning') || msg.includes('warning') || msg.includes('Skipping')) {
          addLog(msg, 'warning');
        } else if (msg.includes('Failed') || msg.includes('failed') || msg.includes('Error') || msg.includes('❌')) {
          addLog(msg, 'danger');
        } else if (msg.includes('Retry')) {
          addLog(msg, 'info');
        } else {
          addLog(msg, 'general');
        }
      }
    );

    const authenticated = await client.authenticate();
    if (!authenticated) {
      addLog('Authentication failed. Terminating upload.', 'danger');
      setIsRunning(false);
      setIsStopping(false);
      return;
    }

    addLog('Validating target configuration IDs...', 'info');
    try {
      const details = await client.fetchConfigDetails(ticketClientId, ticketTypeId, ticketPriorityId);
      addLog(`🎯 Target Client: ${details.clientName} (ID: ${ticketClientId})`, 'success');
      addLog(`🎯 Ticket Type: ${details.ticketTypeName} (ID: ${ticketTypeId})`, 'success');
      addLog(`🎯 Priority: ${details.priorityName} (ID: ${ticketPriorityId})`, 'success');
    } catch (err: any) {
      addLog(`⚠️ Validation check failed: ${err.message}`, 'warning');
    }

    let successCount = 0;
    let failedCount = 0;

    for (let i = 0; i < ticketRows.length; i++) {
      if (stopRequestedRef.current) {
        addLog(`Upload stopped. ${ticketRows.length - i} row(s) left unprocessed.`, 'warning');
        break;
      }

      const row = ticketRows[i];
      const normalisedRow: Record<string, any> = {};
      Object.keys(row).forEach(key => {
        normalisedRow[key.trim().toLowerCase()] = row[key];
      });
      const subject = normalisedRow['subject'] || normalisedRow['subject '] || 'Auto Ticket';

      addLog(`Processing row ${i + 1}/${ticketRows.length}: "${subject}"...`, 'general');
      
      try {
        const success = await client.createTicket(
          row,
          ticketClientId,
          ticketTypeId,
          ticketPriorityId,
          isDryRun
        );
        if (success) {
          successCount++;
        } else {
          failedCount++;
        }
      } catch (err: any) {
        addLog(`Exception creating ticket in row ${i + 1}: ${err.message}`, 'danger');
        failedCount++;
      }

      const current = i + 1;
      setStats({
        total: ticketRows.length,
        current,
        success: successCount,
        failed: failedCount,
      });
      setProgress(Math.round((current / ticketRows.length) * 100));
    }

    const uploadStopped = stopRequestedRef.current;
    addLog(`\n--- Ticket Upload Process ${uploadStopped ? 'Stopped' : 'Completed'} ---`, uploadStopped ? 'warning' : 'info');
    addLog(`Total Processed: ${ticketRows.length}`, 'general');
    addLog(`${isDryRun ? 'Dry Run Simulated' : 'Successfully Created'}: ${successCount}`, 'success');
    addLog(`Failed: ${failedCount}`, 'danger');

    saveAuditLog('upload', `Tickets: ${ticketFile?.name || 'Excel/CSV file'}${isDryRun ? ' (Dry Run)' : ''}`, ticketRows.length, successCount, failedCount);
    stopRequestedRef.current = false;
    setIsRunning(false);
    setIsStopping(false);
  };

  const startBulkDelete = async () => {
    stopRequestedRef.current = false;
    setIsStopping(false);
    setIsRunning(true);
    setProgress(0);
    clearConsole();

    addLog(`Starting bulk delete sequence in mode: ${deleteMode.toUpperCase()}...`, 'info');
    if (dryRun) {
      addLog('DRY RUN IS ENABLED. No actual API deletions will occur.', 'warning');
    }

    const client = new MerlinClient(
      {
        apiUrl: deleteApiUrl,
        username: deleteUsername,
        password: deletePassword,
      },
      (msg) => {
        if (msg.includes('Deleted:') || msg.includes('would delete')) {
          addLog(msg, 'success');
        } else if (msg.includes('Skipping')) {
          addLog(msg, 'warning');
        } else if (msg.includes('Failed')) {
          addLog(msg, 'danger');
        } else {
          addLog(msg, 'general');
        }
      }
    );

    const authenticated = await client.authenticate();
    if (!authenticated) {
      addLog('Authentication failed. Terminating delete.', 'danger');
      setIsRunning(false);
      setIsStopping(false);
      return;
    }

    let successCount = 0;
    let failedCount = 0;
    let attemptedCount = 0;

    const deleteWorkOrdersConcurrently = async (workOrders: any[], initialFailedCount: number = 0) => {
      let nextIndex = 0;
      let completedCount = 0;
      let deleteSuccessCount = 0;
      let deleteFailedCount = 0;
      const workerCount = Math.min(MAX_DELETE_WORKERS, workOrders.length);

      if (workOrders.length === 0) {
        return { success: 0, failed: 0 };
      }

      addLog(`Deleting with up to ${workerCount} concurrent requests...`, 'info');
      setStats({ total: workOrders.length, current: 0, success: 0, failed: initialFailedCount });
      setProgress(0);

      const worker = async () => {
        while (!stopRequestedRef.current && nextIndex < workOrders.length) {
          const currentIndex = nextIndex;
          nextIndex++;

          const success = await client.deleteWorkOrder(workOrders[currentIndex], dryRun);
          completedCount++;

          if (success) {
            deleteSuccessCount++;
          } else {
            deleteFailedCount++;
          }

          setStats({
            total: workOrders.length,
            current: completedCount,
            success: deleteSuccessCount,
            failed: initialFailedCount + deleteFailedCount,
          });
          setProgress(Math.round((completedCount / workOrders.length) * 100));
        }
      };

      await Promise.all(Array.from({ length: workerCount }, () => worker()));
      if (stopRequestedRef.current && completedCount < workOrders.length) {
        addLog(`Delete stopped. ${workOrders.length - completedCount} queued work order(s) were not deleted.`, 'warning');
      }
      return { success: deleteSuccessCount, failed: deleteFailedCount };
    };

    if (deleteMode === 'client_all') {
      const clientName = deleteClientName.trim();
      let clientId = deleteClientId.trim();

      if (!clientId && clientName) {
        addLog(`Searching client by name: "${clientName}"...`, 'info');
        const foundId = await client.getClientIdByName(clientName);
        if (foundId) {
          clientId = foundId;
        }
      }

      if (!clientId) {
        addLog('Client ID or valid Client Name is required.', 'danger');
        setIsRunning(false);
        setIsStopping(false);
        return;
      }

      const locationId = selectedDeleteLocationId.trim();
      if (locationId && locationsClientId !== clientId) {
        addLog('Selected site locations are for a different Client ID. Load locations again.', 'danger');
        setIsRunning(false);
        setIsStopping(false);
        return;
      }

      const selectedLocation = deleteLocations.find((loc) => loc.id === locationId);
      addLog(`Fetching up to ${deleteLimit} work orders for client: ${clientId}${selectedLocation ? ` at location: ${selectedLocation.name}` : ''}...`, 'info');
      const workOrders = await client.fetchWorkOrders(clientId, deleteLimit, deleteOnlyOverdue, locationId || undefined);
      
      if (workOrders.length === 0) {
        addLog('No work orders found matching criteria.', 'warning');
        setIsRunning(false);
        setIsStopping(false);
        return;
      }

      addLog(`Found ${workOrders.length} work orders to delete. Starting loop...`, 'info');
      attemptedCount = workOrders.length;
      const result = await deleteWorkOrdersConcurrently(workOrders);
      successCount = result.success;
      failedCount = result.failed;
    } else {
      // deleteMode === 'sheet'
      if (deleteRows.length === 0) {
        addLog('No rows to process. Please upload a delete Excel/CSV file first.', 'warning');
        setIsRunning(false);
        setIsStopping(false);
        return;
      }

      const totalToDelete = Math.min(deleteLimit, deleteRows.length);
      addLog(`Processing up to ${totalToDelete} deletions from sheet...`, 'info');
      setStats({ total: totalToDelete, current: 0, success: 0, failed: 0 });
      const workOrdersToDelete: any[] = [];

      for (let i = 0; i < totalToDelete; i++) {
        if (stopRequestedRef.current) {
          addLog(`Delete stopped during sheet lookup. ${totalToDelete - i} row(s) left unchecked.`, 'warning');
          break;
        }

        const row = deleteRows[i];
        
        // Accepted column mapping (case-insensitive)
        const rowKeys = Object.keys(row);
        const getRowVal = (keys: string[]) => {
          const matched = rowKeys.find(rk => keys.includes(rk.toLowerCase().trim().replace(/[-_]/g, ' ')));
          return matched ? String(row[matched]).trim() : '';
        };

        const clientId = getRowVal(['org id', 'client id', 'client_id', 'org_id']);
        const woNum = getRowVal(['activity number', 'work order number', 'work_order_number']);
        const locationVal = getRowVal(['location', 'location id', 'location code', 'location name', 'location_id', 'location_code', 'location_name']);

        if (!clientId || !woNum) {
          addLog(`Skipping row ${i + 2}: Client ID and Work Order Number are required.`, 'warning');
          failedCount++;
          continue;
        }

        addLog(`Searching work order ${woNum} for client ${clientId}...`, 'general');
        const workOrder = await client.findWorkOrderByNumber(clientId, woNum);

        if (!workOrder) {
          addLog(`Skipping row ${i + 2}: Work Order ${woNum} not found.`, 'warning');
          failedCount++;
          continue;
        }

        // Verify Location
        if (verifyLocation && locationVal) {
          const locMatched = await client.locationMatches(clientId, workOrder, locationVal);
          if (!locMatched) {
            addLog(`Skipping row ${i + 2}: site location mismatch for WO ${woNum}. Expected: "${locationVal}"`, 'warning');
            failedCount++;
            continue;
          }
        }

        // Verify Overdue
        if (deleteOnlyOverdue && workOrder.is_overdue !== true) {
          addLog(`Skipping row ${i + 2}: WO ${woNum} is not overdue.`, 'warning');
          failedCount++;
          continue;
        }

        workOrdersToDelete.push(workOrder);

        setStats({
          total: totalToDelete,
          current: i + 1,
          success: successCount,
          failed: failedCount,
        });
        setProgress(Math.round(((i + 1) / totalToDelete) * 100));
      }

      attemptedCount = totalToDelete;

      if (workOrdersToDelete.length === 0) {
        addLog('No matched work orders passed validation for deletion.', 'warning');
      } else {
        const result = await deleteWorkOrdersConcurrently(workOrdersToDelete, failedCount);
        successCount = result.success;
        failedCount += result.failed;
      }
    }

    const deleteStopped = stopRequestedRef.current;
    addLog(`\n--- Delete Process ${deleteStopped ? 'Stopped' : 'Completed'} ---`, deleteStopped ? 'warning' : 'info');
    addLog(`Total Attempted: ${attemptedCount}`, 'general');
    addLog(`${dryRun ? 'Simulated' : 'Actual'} Deletions: ${successCount}`, 'success');
    addLog(`Failed/Skipped: ${failedCount}`, 'danger');

    saveAuditLog(
      'delete',
      deleteMode === 'sheet' ? (deleteFile?.name || 'Delete sheet') : `Client: ${deleteClientName || deleteClientId}${selectedDeleteLocationId ? ` | Location: ${deleteLocations.find((loc) => loc.id === selectedDeleteLocationId)?.name || selectedDeleteLocationId}` : ''}`,
      attemptedCount,
      successCount,
      failedCount
    );
    stopRequestedRef.current = false;
    setIsRunning(false);
    setIsStopping(false);
  };

  const filteredLogs = auditLogs.filter(
    (log) =>
      log.user.toLowerCase().includes(searchQuery.toLowerCase()) ||
      log.source.toLowerCase().includes(searchQuery.toLowerCase()) ||
      log.action.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const clearAuditLogs = () => {
    if (confirm('Are you sure you want to clear the audit logs? This cannot be undone.')) {
      setAuditLogs([]);
      localStorage.removeItem('soteria_audit_logs');
    }
  };

  return (
    <div className="dashboard-wrapper">
      {/* Sidebar Navigation */}
      <aside className="sidebar">
        <div className="sidebar-top">
          <div className="brand">
            <div className="brand-icon">
              <Database size={20} />
            </div>
            <span className="brand-name">Soteria WO</span>
          </div>

          <ul className="nav-menu">
            <li className="nav-item">
              <button
                className={`nav-link transition-all ${activeTab === 'upload' ? 'active' : ''}`}
                onClick={() => { setActiveTab('upload'); clearConsole(); }}
                disabled={isRunning}
              >
                <Upload size={18} />
                <span>Bulk WO Upload</span>
              </button>
            </li>
            <li className="nav-item">
              <button
                className={`nav-link transition-all ${activeTab === 'ticket_upload' ? 'active' : ''}`}
                onClick={() => { setActiveTab('ticket_upload'); clearConsole(); }}
                disabled={isRunning}
              >
                <Upload size={18} />
                <span>Bulk Ticket Upload</span>
              </button>
            </li>
            <li className="nav-item">
              <button
                className={`nav-link transition-all ${activeTab === 'delete' ? 'active' : ''}`}
                onClick={() => { setActiveTab('delete'); clearConsole(); }}
                disabled={isRunning}
              >
                <Trash2 size={18} />
                <span>Bulk Delete</span>
              </button>
            </li>
            <li className="nav-item">
              <button
                className={`nav-link transition-all ${activeTab === 'logs' ? 'active' : ''}`}
                onClick={() => setActiveTab('logs')}
                disabled={isRunning}
              >
                <FileText size={18} />
                <span>Audit Logs</span>
              </button>
            </li>
          </ul>
        </div>

        <div className="sidebar-bottom">
          <div className="user-info">
            <div className="user-avatar">
              <User size={14} />
            </div>
            <div className="user-details">
              <span className="user-role">Administrator</span>
              <span className="user-email" title={userEmail}>{userEmail}</span>
            </div>
          </div>
          <button className="logout-btn transition-all" onClick={onLogout} disabled={isRunning}>
            <LogOut size={16} />
            <span>Logout</span>
          </button>
        </div>
      </aside>

      {/* Main dashboard content */}
      <main className="main-content">
        <header className="page-header">
          <div>
            <h1 className="page-title">
              {activeTab === 'upload' && 'Bulk Work Order Upload'}
              {activeTab === 'ticket_upload' && 'Bulk Ticket Upload'}
              {activeTab === 'delete' && 'Bulk Work Order Delete'}
              {activeTab === 'logs' && 'Audit Log Trail'}
            </h1>
            <p className="page-subtitle">
              {activeTab === 'upload' && 'Upload Excel sheets to create work orders in bulk across clients.'}
              {activeTab === 'ticket_upload' && 'Upload Excel sheets to create tickets in bulk.'}
              {activeTab === 'delete' && 'Clean up work orders in bulk from database search or excel sheets.'}
              {activeTab === 'logs' && 'Historical records of all uploads and deletes processed.'}
            </p>
          </div>
        </header>

        {activeTab !== 'logs' ? (
          <div className="dashboard-grid">
            {/* Left Column: Configurations & Files */}
            <div className="glass-panel animated-slide-up" style={{ animationDelay: '0.1s' }}>
              {activeTab === 'upload' && (
                <>
                  <div className="panel-header">
                    <Settings size={18} />
                    <span>Upload Configurations</span>
                  </div>

                  <div className="form-grid">
                    <div className="input-group">
                      <label className="input-label">APP URL</label>
                      <input
                        type="text"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={uploadApiUrl}
                        onChange={(e) => setUploadApiUrl(e.target.value)}
                        disabled={isRunning}
                        readOnly
                      />
                    </div>
                  </div>

                  <div className="form-grid two-col">
                    <div className="input-group">
                      <label className="input-label">USERNAME</label>
                      <input
                        type="text"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={uploadUsername}
                        onChange={(e) => setUploadUsername(e.target.value)}
                        disabled={isRunning}
                      />
                    </div>
                    <div className="input-group">
                      <label className="input-label">PASSWORD</label>
                      <input
                        type="password"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={uploadPassword}
                        onChange={(e) => setUploadPassword(e.target.value)}
                        disabled={isRunning}
                      />
                    </div>
                  </div>

                  <div className="panel-header" style={{ marginTop: '24px', marginBottom: '16px' }}>
                    <FileSpreadsheet size={18} />
                    <span>Excel or CSV Sheet File</span>
                  </div>

                  <div className={`file-uploader ${uploadFile ? 'has-file' : ''}`}>
                    <div className="file-uploader-icon">
                      <Upload size={20} />
                    </div>
                    <div className="file-info">
                      {uploadFile ? (
                        <div>
                          File selected: <span className="file-name">{uploadFile.name}</span>
                          <div style={{ marginTop: '4px', fontSize: '11px', color: 'var(--text-muted)' }}>
                            Rows detected: {uploadRows.length}
                          </div>
                        </div>
                      ) : (
                        <span>Drag and drop Excel/CSV or Click to Browse</span>
                      )}
                    </div>
                    <input
                      type="file"
                      className="file-uploader-input"
                      accept=".xlsx, .xls, .csv"
                      onChange={handleUploadFileChange}
                      disabled={isRunning}
                    />
                  </div>

                  <div className="action-row">
                    <button
                      className="action-btn transition-all"
                      onClick={startBulkUpload}
                      disabled={isRunning || uploadRows.length === 0}
                    >
                      <Play size={16} />
                      <span>Start Bulk Upload</span>
                    </button>
                    {isRunning && (
                      <button
                        className="action-btn stop-btn transition-all"
                        onClick={requestStop}
                        disabled={isStopping}
                      >
                        <X size={16} />
                        <span>{isStopping ? 'Stopping' : 'Stop'}</span>
                      </button>
                    )}
                  </div>
                </>
              )}

              {activeTab === 'ticket_upload' && (
                <>
                  <div className="panel-header">
                    <Settings size={18} />
                    <span>Ticket Upload Configurations</span>
                  </div>

                  <div className="form-grid">
                    <div className="input-group">
                      <label className="input-label">API URL</label>
                      <input
                        type="text"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={ticketApiUrl}
                        onChange={(e) => setTicketApiUrl(e.target.value)}
                        disabled={isRunning}
                      />
                    </div>
                  </div>

                  <div className="form-grid two-col">
                    <div className="input-group">
                      <label className="input-label">EMAIL</label>
                      <input
                        type="text"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={ticketEmail}
                        onChange={(e) => setTicketEmail(e.target.value)}
                        disabled={isRunning}
                      />
                    </div>
                    <div className="input-group">
                      <label className="input-label">PASSWORD</label>
                      <input
                        type="password"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={ticketPassword}
                        onChange={(e) => setTicketPassword(e.target.value)}
                        disabled={isRunning}
                      />
                    </div>
                  </div>

                  <div className="form-grid two-col">
                    <div className="input-group">
                      <label className="input-label">CLIENT ID</label>
                      <input
                        type="text"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={ticketClientId}
                        onChange={(e) => setTicketClientId(e.target.value)}
                        disabled={isRunning}
                      />
                    </div>
                    <div className="input-group">
                      <label className="input-label">TICKET TYPE ID</label>
                      <input
                        type="text"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={ticketTypeId}
                        onChange={(e) => setTicketTypeId(e.target.value)}
                        disabled={isRunning}
                      />
                    </div>
                  </div>

                  <div className="form-grid">
                    <div className="input-group">
                      <label className="input-label">PRIORITY ID</label>
                      <input
                        type="text"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={ticketPriorityId}
                        onChange={(e) => setTicketPriorityId(e.target.value)}
                        disabled={isRunning}
                      />
                    </div>
                  </div>

                  <div className="panel-header" style={{ marginTop: '24px', marginBottom: '16px' }}>
                    <FileSpreadsheet size={18} />
                    <span>Tickets Excel or CSV Sheet File</span>
                  </div>

                  <div className={`file-uploader ${ticketFile ? 'has-file' : ''}`}>
                    <div className="file-uploader-icon">
                      <Upload size={20} />
                    </div>
                    <div className="file-info">
                      {ticketFile ? (
                        <div>
                          File selected: <span className="file-name">{ticketFile.name}</span>
                          <div style={{ marginTop: '4px', fontSize: '11px', color: 'var(--text-muted)' }}>
                            Rows detected: {ticketRows.length}
                          </div>
                        </div>
                      ) : (
                        <span>Drag and drop Excel/CSV or Click to Browse</span>
                      )}
                    </div>
                    <input
                      type="file"
                      className="file-uploader-input"
                      accept=".xlsx, .xls, .csv"
                      onChange={handleTicketFileChange}
                      disabled={isRunning}
                    />
                  </div>

                  <div className="action-row" style={{ display: 'flex', gap: '10px' }}>
                    <button
                      className="action-btn secondary transition-all"
                      onClick={() => startTicketBulkUpload(true)}
                      disabled={isRunning || ticketRows.length === 0}
                    >
                      <Play size={16} />
                      <span>Dry Run Test</span>
                    </button>
                    <button
                      className="action-btn transition-all"
                      onClick={() => startTicketBulkUpload(false)}
                      disabled={isRunning || ticketRows.length === 0}
                    >
                      <Play size={16} />
                      <span>Start Bulk Upload</span>
                    </button>
                    {isRunning && (
                      <button
                        className="action-btn stop-btn transition-all"
                        onClick={requestStop}
                        disabled={isStopping}
                      >
                        <X size={16} />
                        <span>{isStopping ? 'Stopping' : 'Stop'}</span>
                      </button>
                    )}
                  </div>
                </>
              )}

              {activeTab === 'delete' && (
                <>
                  <div className="panel-header">
                    <Settings size={18} />
                    <span>Delete Configurations</span>
                  </div>

                  <div className="form-grid">
                    <div className="input-group">
                      <label className="input-label">APP URL</label>
                      <input
                        type="text"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={deleteApiUrl}
                        onChange={(e) => setDeleteApiUrl(e.target.value)}
                        disabled={isRunning}
                      />
                    </div>
                  </div>

                  <div className="form-grid two-col">
                    <div className="input-group">
                      <label className="input-label">USERNAME</label>
                      <input
                        type="text"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={deleteUsername}
                        onChange={(e) => setDeleteUsername(e.target.value)}
                        disabled={isRunning}
                      />
                    </div>
                    <div className="input-group">
                      <label className="input-label">PASSWORD</label>
                      <input
                        type="password"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={deletePassword}
                        onChange={(e) => setDeletePassword(e.target.value)}
                        disabled={isRunning}
                      />
                    </div>
                  </div>

                  <div className="form-grid two-col" style={{ marginTop: '8px' }}>
                    <div className="input-group">
                      <label className="input-label">DELETE MODE</label>
                      <select
                        className="input-field"
                        style={{ padding: '14px' }}
                        value={deleteMode}
                        onChange={(e) => setDeleteMode(e.target.value as 'sheet' | 'client_all')}
                        disabled={isRunning}
                      >
                        <option value="sheet">CSV/Excel Sheet file</option>
                        <option value="client_all">All for Client ID/Name</option>
                      </select>
                    </div>

                    <div className="input-group">
                      <label className="input-label">LIMIT (Max items to delete)</label>
                      <input
                        type="number"
                        className="input-field"
                        style={{ paddingLeft: '14px' }}
                        value={deleteLimit}
                        onChange={(e) => setDeleteLimit(parseInt(e.target.value) || 0)}
                        disabled={isRunning}
                        min={1}
                        max={3300}
                      />
                    </div>
                  </div>

                  {deleteMode === 'client_all' ? (
                    <div className="form-grid two-col" style={{ marginTop: '8px' }}>
                      <div className="input-group">
                        <label className="input-label">CLIENT NAME</label>
                        <input
                          type="text"
                          className="input-field"
                          style={{ paddingLeft: '14px' }}
                          placeholder="e.g. CBRE"
                          value={deleteClientName}
                          onChange={(e) => setDeleteClientName(e.target.value)}
                          disabled={isRunning}
                        />
                      </div>
                      <div className="input-group">
                        <label className="input-label">CLIENT ID (Overrides Name)</label>
                        <input
                          type="text"
                          className="input-field"
                          style={{ paddingLeft: '14px' }}
                          placeholder="UUID (Safer)"
                          value={deleteClientId}
                          onChange={(e) => {
                            setDeleteClientId(e.target.value);
                            resetDeleteLocations();
                          }}
                          disabled={isRunning}
                        />
                      </div>
                      <div className="input-group">
                        <label className="input-label">SITE LOCATION</label>
                        <select
                          className="input-field"
                          style={{ padding: '14px' }}
                          value={selectedDeleteLocationId}
                          onChange={(e) => setSelectedDeleteLocationId(e.target.value)}
                          disabled={isRunning || isLoadingLocations || deleteLocations.length === 0}
                        >
                          <option value="">
                            {deleteLocations.length === 0 ? 'Load locations first' : 'All site locations'}
                          </option>
                          {deleteLocations.map((location) => (
                            <option key={location.id} value={location.id}>
                              {location.code ? `${location.name} (${location.code})` : location.name}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div className="input-group">
                        <label className="input-label">&nbsp;</label>
                        <button
                          type="button"
                          className="action-btn secondary transition-all"
                          onClick={loadDeleteLocations}
                          disabled={isRunning || isLoadingLocations || !deleteClientId.trim()}
                        >
                          {isLoadingLocations ? (
                            <span className="loading-spinner"></span>
                          ) : (
                            <MapPin size={16} />
                          )}
                          <span>{isLoadingLocations ? 'Loading Locations' : 'Load Locations'}</span>
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div style={{ marginTop: '16px' }}>
                      <div className="panel-header" style={{ marginBottom: '16px' }}>
                        <FileSpreadsheet size={18} />
                        <span>Delete Target Sheet</span>
                      </div>

                      <div className={`file-uploader ${deleteFile ? 'has-file' : ''}`}>
                        <div className="file-uploader-icon">
                          <Upload size={20} />
                        </div>
                        <div className="file-info">
                          {deleteFile ? (
                            <div>
                              File selected: <span className="file-name">{deleteFile.name}</span>
                              <div style={{ marginTop: '4px', fontSize: '11px', color: 'var(--text-muted)' }}>
                                Rows detected: {deleteRows.length}
                              </div>
                            </div>
                          ) : (
                            <span>Drag and drop delete sheet (.csv, .xlsx) or Click to Browse</span>
                          )}
                        </div>
                        <input
                          type="file"
                          className="file-uploader-input"
                          accept=".xlsx, .xls, .csv"
                          onChange={handleDeleteFileChange}
                          disabled={isRunning}
                        />
                      </div>
                    </div>
                  )}

                  {/* Toggle controls */}
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '16px', marginBottom: '24px' }}>
                    <div className="toggle-switch-container">
                      <div className="toggle-label-wrap">
                        <span className="toggle-title">DRY RUN ONLY</span>
                        <span className="toggle-desc">Logs matching entries without deleting.</span>
                      </div>
                      <label className="switch">
                        <input
                          type="checkbox"
                          checked={dryRun}
                          onChange={(e) => setDryRun(e.target.checked)}
                          disabled={isRunning}
                        />
                        <span className="slider"></span>
                      </label>
                    </div>

                    <div className="toggle-switch-container">
                      <div className="toggle-label-wrap">
                        <span className="toggle-title">ONLY OVERDUE WORKORDERS</span>
                        <span className="toggle-desc">Checks if the workorder is overdue.</span>
                      </div>
                      <label className="switch">
                        <input
                          type="checkbox"
                          checked={deleteOnlyOverdue}
                          onChange={(e) => setDeleteOnlyOverdue(e.target.checked)}
                          disabled={isRunning}
                        />
                        <span className="slider"></span>
                      </label>
                    </div>

                    {deleteMode === 'sheet' && (
                      <div className="toggle-switch-container">
                        <div className="toggle-label-wrap">
                          <span className="toggle-title">VERIFY SITE LOCATION</span>
                          <span className="toggle-desc">Skips deletions if locations mismatch.</span>
                        </div>
                        <label className="switch">
                          <input
                            type="checkbox"
                            checked={verifyLocation}
                            onChange={(e) => setVerifyLocation(e.target.checked)}
                            disabled={isRunning}
                          />
                          <span className="slider"></span>
                        </label>
                      </div>
                    )}
                  </div>

                  <div className="action-row">
                    <button
                      className="action-btn danger-btn transition-all"
                      onClick={startBulkDelete}
                      disabled={isRunning || isLoadingLocations || (deleteMode === 'sheet' && deleteRows.length === 0)}
                    >
                      <Trash2 size={16} />
                      <span>{dryRun ? 'Simulate Deletes' : 'Perform Bulk Delete'}</span>
                    </button>
                    {isRunning && (
                      <button
                        className="action-btn stop-btn transition-all"
                        onClick={requestStop}
                        disabled={isStopping}
                      >
                        <X size={16} />
                        <span>{isStopping ? 'Stopping' : 'Stop'}</span>
                      </button>
                    )}
                  </div>
                </>
              )}
            </div>

            {/* Right Column: Progress & Terminal Logs */}
            <div className="glass-panel console-panel animated-slide-up" style={{ animationDelay: '0.2s' }}>
              <div className="panel-header" style={{ justifyContent: 'space-between', marginBottom: '16px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                  <Terminal size={18} />
                  <span>Execution Output Logs</span>
                </div>
                {isRunning && <div className="loading-spinner"></div>}
              </div>

              {/* Stats Strip */}
              <div className="stats-strip">
                <div className="stat-card">
                  <div className="stat-label">Processed</div>
                  <div className="stat-value primary">
                    {stats.current} / {stats.total}
                  </div>
                </div>
                <div className="stat-card">
                  <div className="stat-label">Successful</div>
                  <div className="stat-value success">{stats.success}</div>
                </div>
                <div className="stat-card">
                  <div className="stat-label">Errors</div>
                  <div className="stat-value danger">{stats.failed}</div>
                </div>
              </div>

              {/* Progress meter */}
              <div className="progress-bar-container">
                <div className="progress-bar" style={{ width: `${progress}%` }}></div>
              </div>

              {/* Console terminal logs */}
              <div className="console-actions">
                <button className="console-clear-btn" onClick={clearConsole} disabled={isRunning}>
                  <X size={12} />
                  <span>Clear Console</span>
                </button>
              </div>

              <div className="console-box" ref={consoleBoxRef}>
                {logs.length === 0 ? (
                  <div className="log-line log-debug">Console idle. Awaiting configuration launch...</div>
                ) : (
                  logs.map((log, index) => (
                    <div
                      key={index}
                      className={`log-line 
                        ${log.type === 'success' ? 'log-success' : ''}
                        ${log.type === 'warning' ? 'log-warning' : ''}
                        ${log.type === 'danger' ? 'log-danger' : ''}
                        ${log.type === 'info' ? 'log-info' : ''}
                        ${log.type === 'debug' ? 'log-debug' : ''}
                        ${log.type === 'general' ? 'log-general' : ''}
                      `}
                    >
                      {log.text}
                    </div>
                  ))
                )}
              </div>
            </div>
          </div>
        ) : (
          /* Logs Tab View */
          <div className="glass-panel table-panel animated-slide-up">
            <div className="table-controls">
              <div className="search-wrap">
                <Search className="search-icon" size={16} />
                <input
                  type="text"
                  className="search-input"
                  placeholder="Search logs by action, file..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                />
              </div>

              <button className="logout-btn transition-all" style={{ width: 'auto', background: 'transparent' }} onClick={clearAuditLogs}>
                <Trash2 size={16} />
                <span>Clear Audit Trail</span>
              </button>
            </div>

            <div className="table-scroll">
              <table className="audit-table">
                <thead>
                  <tr>
                    <th>TIMESTAMP</th>
                    <th>USER</th>
                    <th>ACTION</th>
                    <th>SOURCE / SCOPE</th>
                    <th>TOTAL</th>
                    <th>SUCCESS</th>
                    <th>ERRORS / SKIPPED</th>
                    <th>STATUS</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredLogs.length === 0 ? (
                    <tr>
                      <td colSpan={8}>
                        <div className="empty-state">
                          <div className="empty-icon">
                            <Info size={20} />
                          </div>
                          <span>No execution audit logs found.</span>
                        </div>
                      </td>
                    </tr>
                  ) : (
                    filteredLogs.map((log) => {
                      const dateStr = new Date(log.timestamp).toLocaleString();
                      const successRate = log.total > 0 ? (log.success / log.total) * 100 : 0;
                      return (
                        <tr key={log.id}>
                          <td style={{ whiteSpace: 'nowrap' }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                              <Clock size={12} style={{ color: 'var(--text-muted)' }} />
                              <span>{dateStr}</span>
                            </div>
                          </td>
                          <td style={{ color: 'var(--text-secondary)' }}>{log.user}</td>
                          <td>
                            <span className={`badge ${log.action}`}>
                              {log.action === 'upload' ? 'Upload' : 'Delete'}
                            </span>
                          </td>
                          <td style={{ fontWeight: 500, maxWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={log.source}>
                            {log.source}
                          </td>
                          <td style={{ fontWeight: 600 }}>{log.total}</td>
                          <td style={{ color: 'var(--success)', fontWeight: 600 }}>{log.success}</td>
                          <td style={{ color: log.failed > 0 ? 'var(--danger)' : 'var(--text-muted)', fontWeight: 600 }}>
                            {log.failed}
                          </td>
                          <td>
                            {successRate === 100 ? (
                              <span style={{ color: 'var(--success)', display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '11px', fontWeight: 600 }}>
                                <CheckCircle size={12} /> Complete
                              </span>
                            ) : successRate > 0 ? (
                              <span style={{ color: 'var(--warning)', display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '11px', fontWeight: 600 }}>
                                <AlertTriangle size={12} /> Partial
                              </span>
                            ) : (
                              <span style={{ color: 'var(--danger)', display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '11px', fontWeight: 600 }}>
                                <ShieldAlert size={12} /> Failed
                              </span>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
