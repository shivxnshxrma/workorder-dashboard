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
  ShieldAlert
} from 'lucide-react';
import { MerlinClient } from '@/lib/merlin';

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

export default function DashboardClient({ userEmail, onLogout }: { userEmail: string; onLogout: () => void }) {
  const [activeTab, setActiveTab] = useState<'upload' | 'delete' | 'logs'>('upload');
  
  // App Configs
  const [uploadApiUrl, setUploadApiUrl] = useState('https://api-mneo-cbre.merlin-soteria.in/api/v1/');
  const [uploadUsername, setUploadUsername] = useState('');
  const [uploadPassword, setUploadPassword] = useState('');
  
  const [deleteApiUrl, setDeleteApiUrl] = useState('https://api-merlin.tenonfm-india.com/api/v1/');
  const [deleteUsername, setDeleteUsername] = useState('');
  const [deletePassword, setDeletePassword] = useState('');
  
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
  
  // Process State
  const [isRunning, setIsRunning] = useState(false);
  const [progress, setProgress] = useState(0);
  const [stats, setStats] = useState({ total: 0, current: 0, success: 0, failed: 0 });
  const [logs, setLogs] = useState<LogLine[]>([]);
  
  // Audit Logs State
  const [auditLogs, setAuditLogs] = useState<AuditLog[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  
  const consoleEndRef = useRef<HTMLDivElement>(null);

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

  // Auto-scroll console
  useEffect(() => {
    if (consoleEndRef.current) {
      consoleEndRef.current.scrollIntoView({ behavior: 'smooth' });
    }
  }, [logs]);

  const addLog = (text: string, type: LogLine['type'] = 'general') => {
    const timestamp = new Date().toLocaleTimeString();
    setLogs((prev) => [...prev, { type, text: `[${timestamp}] ${text}` }]);
  };

  const clearConsole = () => {
    setLogs([]);
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

  const startBulkUpload = async () => {
    if (uploadRows.length === 0) {
      addLog('No rows to upload. Please upload an Excel/CSV file first.', 'warning');
      return;
    }

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
      return;
    }

    await client.preloadGeneralCache();

    // Group client caching
    const uniqueClients = Array.from(new Set(uploadRows.map((r: any) => String(r.Client || '').trim()).filter(Boolean)));
    for (const clientName of uniqueClients) {
      const clientId = await client.getClientIdByName(clientName);
      if (clientId) {
        await client.preloadCacheForClient(clientId, clientName);
      }
    }

    let successCount = 0;
    let failedCount = 0;

    for (let i = 0; i < uploadRows.length; i++) {
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

    addLog('\n--- Upload Process Completed ---', 'info');
    addLog(`Total Processed: ${uploadRows.length}`, 'general');
    addLog(`Successfully Created: ${successCount}`, 'success');
    addLog(`Failed: ${failedCount}`, 'danger');

    saveAuditLog('upload', uploadFile?.name || 'Excel/CSV file', uploadRows.length, successCount, failedCount);
    setIsRunning(false);
  };

  const startBulkDelete = async () => {
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
      return;
    }

    let successCount = 0;
    let failedCount = 0;
    let attemptedCount = 0;

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
        return;
      }

      addLog(`Fetching up to ${deleteLimit} work orders for client: ${clientId}...`, 'info');
      const workOrders = await client.fetchWorkOrders(clientId, deleteLimit, deleteOnlyOverdue);
      
      if (workOrders.length === 0) {
        addLog('No work orders found matching criteria.', 'warning');
        setIsRunning(false);
        return;
      }

      addLog(`Found ${workOrders.length} work orders to delete. Starting loop...`, 'info');
      setStats({ total: workOrders.length, current: 0, success: 0, failed: 0 });

      for (let i = 0; i < workOrders.length; i++) {
        attemptedCount++;
        const wo = workOrders[i];
        const success = await client.deleteWorkOrder(wo, dryRun);
        if (success) {
          successCount++;
        } else {
          failedCount++;
        }

        setStats({
          total: workOrders.length,
          current: i + 1,
          success: successCount,
          failed: failedCount,
        });
        setProgress(Math.round(((i + 1) / workOrders.length) * 100));
      }
    } else {
      // deleteMode === 'sheet'
      if (deleteRows.length === 0) {
        addLog('No rows to process. Please upload a delete Excel/CSV file first.', 'warning');
        setIsRunning(false);
        return;
      }

      const totalToDelete = Math.min(deleteLimit, deleteRows.length);
      addLog(`Processing up to ${totalToDelete} deletions from sheet...`, 'info');
      setStats({ total: totalToDelete, current: 0, success: 0, failed: 0 });

      for (let i = 0; i < totalToDelete; i++) {
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

        attemptedCount++;
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

        const success = await client.deleteWorkOrder(workOrder, dryRun);
        if (success) {
          successCount++;
        } else {
          failedCount++;
        }

        setStats({
          total: totalToDelete,
          current: attemptedCount,
          success: successCount,
          failed: failedCount,
        });
        setProgress(Math.round((attemptedCount / totalToDelete) * 100));
      }
    }

    addLog('\n--- Delete Process Completed ---', 'info');
    addLog(`Total Attempted: ${attemptedCount}`, 'general');
    addLog(`${dryRun ? 'Simulated' : 'Actual'} Deletions: ${successCount}`, 'success');
    addLog(`Failed/Skipped: ${failedCount}`, 'danger');

    saveAuditLog(
      'delete',
      deleteMode === 'sheet' ? (deleteFile?.name || 'Delete sheet') : `All for client: ${deleteClientName || deleteClientId}`,
      attemptedCount,
      successCount,
      failedCount
    );
    setIsRunning(false);
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
                <span>Bulk Upload</span>
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
              {activeTab === 'delete' && 'Bulk Work Order Delete'}
              {activeTab === 'logs' && 'Audit Log Trail'}
            </h1>
            <p className="page-subtitle">
              {activeTab === 'upload' && 'Upload Excel sheets to create work orders in bulk across clients.'}
              {activeTab === 'delete' && 'Clean up work orders in bulk from database search or excel sheets.'}
              {activeTab === 'logs' && 'Historical records of all uploads and deletes processed.'}
            </p>
          </div>
        </header>

        {activeTab !== 'logs' ? (
          <div className="dashboard-grid">
            {/* Left Column: Configurations & Files */}
            <div className="glass-panel animated-slide-up" style={{ animationDelay: '0.1s' }}>
              {activeTab === 'upload' ? (
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

                  <button
                    className="action-btn transition-all"
                    onClick={startBulkUpload}
                    disabled={isRunning || uploadRows.length === 0}
                  >
                    <Play size={16} />
                    <span>Start Bulk Upload</span>
                  </button>
                </>
              ) : (
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
                        onChange={(e: any) => setDeleteMode(e.target.value)}
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
                          onChange={(e) => setDeleteClientId(e.target.value)}
                          disabled={isRunning}
                        />
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

                  <button
                    className="action-btn danger-btn transition-all"
                    onClick={startBulkDelete}
                    disabled={isRunning || (deleteMode === 'sheet' && deleteRows.length === 0)}
                  >
                    <Trash2 size={16} />
                    <span>{dryRun ? 'Simulate Deletes' : 'Perform Bulk Delete'}</span>
                  </button>
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

              <div className="console-box">
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
                <div ref={consoleEndRef} />
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
