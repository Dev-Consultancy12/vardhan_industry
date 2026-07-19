import React, { useState, useEffect } from 'react';
import { RefreshCcw, Save, Search, CheckCircle, Edit3, Upload } from 'lucide-react';
import './TestReadingsEditor.css'; 

const formatValue = (val) => {
    if (typeof val === 'number') {
        // Fix floating point issues like 0.41000000000000003
        return parseFloat(val.toPrecision(10));
    }
    return val;
};

const TestReadingsEditor = () => {
  const [data, setData] = useState({});
  const [groups, setGroups] = useState([]);
  const [selectedGroup, setSelectedGroup] = useState('');
  const [status, setStatus] = useState('loading'); // loading, success, error
  const [errorMsg, setErrorMsg] = useState('');
  const [saving, setSaving] = useState(false);
  
  // Track changes: [{row, col, val}]
  const [pendingChanges, setPendingChanges] = useState([]);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    setStatus('loading');
    try {
      const apiUrl = import.meta.env.VITE_API_URL || `http://${window.location.hostname}:8090`;
      const response = await fetch(`${apiUrl}/api/test-readings`);
      if (!response.ok) throw new Error('Failed to fetch test readings');
      const json = await response.json();
      setData(json);
      const grps = Object.keys(json);
      setGroups(grps);
      if (grps.length > 0) setSelectedGroup(grps[0]);
      setStatus('success');
      setPendingChanges([]);
    } catch (err) {
      setErrorMsg(err.message);
      setStatus('error');
    }
  };

  const handleCellEdit = (paramName, type, index, row, col, currentVal) => {
    const newValStr = prompt(`Edit ${type} for ${paramName}:`, currentVal);
    if (newValStr === null) return; // cancelled
    
    // Attempt to parse to number if possible, else string
    let newVal = newValStr.trim();
    if (newVal !== '' && !isNaN(newVal)) {
        newVal = Number(newVal);
    }
    
    if (newVal !== currentVal) {
        // Deep copy data to update state
        const newData = JSON.parse(JSON.stringify(data));
        if (type === 'pool') {
            newData[selectedGroup][paramName].pool[index].val = newVal;
        } else if (type === 'param') {
            newData[selectedGroup][paramName].param.val = newVal;
        } else if (type === 'spec') {
            newData[selectedGroup][paramName].spec.val = newVal;
        }
        setData(newData);
        
        // Add to pending changes
        const existingChangeIdx = pendingChanges.findIndex(c => c.row === row && c.col === col);
        const newChanges = [...pendingChanges];
        if (existingChangeIdx >= 0) {
            newChanges[existingChangeIdx].val = newVal;
        } else {
            newChanges.push({ row, col, val: newVal });
        }
        setPendingChanges(newChanges);
    }
  };

  const handleSave = async () => {
    if (pendingChanges.length === 0) return;
    setSaving(true);
    try {
        const apiUrl = import.meta.env.VITE_API_URL || `http://${window.location.hostname}:8090`;
        const response = await fetch(`${apiUrl}/api/test-readings/update`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(pendingChanges)
        });
        
        if (!response.ok) throw new Error('Failed to save updates');
        alert("Successfully saved updates to the Master Test Readings file!");
        setPendingChanges([]);
    } catch (err) {
        alert("Error saving: " + err.message);
    } finally {
        setSaving(false);
    }
  };

  const handleUploadTracker = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    setSaving(true);
    const formData = new FormData();
    formData.append('file', file);
    
    try {
        const apiUrl = import.meta.env.VITE_API_URL || `http://${window.location.hostname}:8090`;
        const response = await fetch(`${apiUrl}/api/upload-tracker`, {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const errData = await response.json();
            throw new Error(errData.detail || 'Failed to upload tracker');
        }
        
        alert("Item Codes Tracker successfully updated!");
    } catch (err) {
        alert("Error uploading tracker: " + err.message);
    } finally {
        setSaving(false);
        e.target.value = ''; // Reset input
    }
  };

  if (status === 'loading') {
      return (
          <div className="glass-panel editor-loading">
              <RefreshCcw size={48} className="spinner" />
              <div className="loading-text">Loading Master Data...</div>
          </div>
      )
  }
  
  if (status === 'error') {
      return (
          <div className="glass-panel editor-loading">
              <div className="loading-text" style={{color: '#ef4444'}}>Error Occurred</div>
              <div className="upload-subtext" style={{color: '#f87171'}}>{errorMsg}</div>
              <button className="reset-btn" onClick={fetchData}>Try Again</button>
          </div>
      )
  }

  const groupData = data[selectedGroup] || {};

  return (
    <div className="glass-panel editor-panel">
      <div className="editor-header">
        <div className="editor-title-area">
          <h2>Master Data Editor</h2>
          <p className="subtitle">Modify the exact base pools for random generation.</p>
        </div>
        
        <div className="editor-actions">
           {pendingChanges.length > 0 && (
               <div className="changes-badge">{pendingChanges.length} unsaved changes</div>
           )}
           <button 
             className={`save-btn ${pendingChanges.length > 0 ? 'active' : ''}`} 
             onClick={handleSave}
             disabled={pendingChanges.length === 0 || saving}
           >
             {saving ? <RefreshCcw size={18} className="spinner" /> : <Save size={18} />}
             Save to Excel
           </button>
        </div>
      </div>
      
      <div className="editor-controls" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div className="group-selector">
              <label>Target Group</label>
              <select value={selectedGroup} onChange={(e) => setSelectedGroup(e.target.value)}>
                  {groups.map(g => (
                      <option key={g} value={g}>{g}</option>
                  ))}
              </select>
          </div>
          <div className="tracker-upload">
              <input 
                  type="file" 
                  id="tracker-upload-input" 
                  accept=".xlsx" 
                  style={{ display: 'none' }}
                  onChange={handleUploadTracker}
              />
              <label htmlFor="tracker-upload-input" className="upload-btn">
                  {saving ? <RefreshCcw size={16} className="spinner" /> : <Upload size={16} />}
                  Upload New Item Codes Tracker
              </label>
          </div>
      </div>

      <div className="table-container">
          <table className="readings-table">
              <thead>
                  <tr>
                      <th>Parameter</th>
                      <th>Specification</th>
                      <th>Data Pool (Click to edit)</th>
                  </tr>
              </thead>
              <tbody>
                  {Object.entries(groupData).map(([paramName, pData]) => (
                      <tr key={paramName}>
                          <td className="param-cell">
                              <div 
                                className="clickable-text"
                                onClick={() => handleCellEdit(paramName, 'param', null, pData.param.row, pData.param.col, pData.param.val)}
                                title="Click to edit Parameter"
                              >
                                {pData.param.val} <Edit3 size={12} className="edit-hint-icon" />
                              </div>
                          </td>
                          <td className="spec-cell">
                              {pData.spec ? (
                                  <div 
                                    className="clickable-text"
                                    onClick={() => handleCellEdit(paramName, 'spec', null, pData.spec.row, pData.spec.col, pData.spec.val)}
                                    title="Click to edit Specification"
                                  >
                                    {pData.spec.val} <Edit3 size={12} className="edit-hint-icon" />
                                  </div>
                              ) : (
                                  <span className="empty-spec">-</span>
                              )}
                          </td>
                          <td className="pool-cell">
                              {pData.pool.length === 0 ? (
                                  <span className="empty-pool">No values mapped (Randomized)</span>
                              ) : (
                                  <div className="pool-chips">
                                      {pData.pool.map((item, idx) => (
                                          <div 
                                            key={`${item.row}-${item.col}`} 
                                            className="pool-chip"
                                            onClick={() => handleCellEdit(paramName, 'pool', idx, item.row, item.col, item.val)}
                                          >
                                              {formatValue(item.val)}
                                          </div>
                                      ))}
                                  </div>
                              )}
                          </td>
                      </tr>
                  ))}
              </tbody>
          </table>
      </div>
    </div>
  );
};

export default TestReadingsEditor;
