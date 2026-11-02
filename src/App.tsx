import { useState, useEffect } from 'react';
import {
  Users, UserPlus, Upload, Trash2, MailPlus,
  CheckCircle2, XCircle, Check, Building2,
  ChevronDown, ChevronRight, Star, History
} from 'lucide-react';
import './index.css';

// Types
interface Person {
  id: string | number;
  name: string;
  title: string;
  email: string;
  department: string;
  reportsTo?: string;
}

type TabType = 'create' | 'use';
type ViewMode = 'department' | 'orgchart';

const loadSavedList = <T,>(key: string): T[] => {
  try {
    const parsed = JSON.parse(localStorage.getItem(key) ?? '[]');
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    console.error(`Failed to parse saved ${key}`);
    return [];
  }
};

const isValidEmail = (value: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);

const sameEmail = (a: string, b?: string) => a.toLowerCase() === b?.toLowerCase();

export default function App() {
  const [activeTab, setActiveTab] = useState<TabType>('use');
  const [viewMode, setViewMode] = useState<ViewMode>('department');
  const [orgData, setOrgData] = useState<Person[]>(() =>
    // Older versions saved Human Resources as HR
    loadSavedList<Person>('orgChartData').map(p => p.department === 'HR' ? { ...p, department: 'Human Resources' } : p)
  );
  const [selectedTo, setSelectedTo] = useState<Person[]>([]);
  const [selectedCc, setSelectedCc] = useState<Person[]>([]);
  const [ccAncestors, setCcAncestors] = useState(true);

  // Form state
  const [name, setName] = useState('');
  const [title, setTitle] = useState('');
  const [email, setEmail] = useState('');
  const [department, setDepartment] = useState('');
  const [reportsTo, setReportsTo] = useState('');
  const [importText, setImportText] = useState('');
  const [searchTerm, setSearchTerm] = useState('');

  // Favorites & Recents state
  const [favorites, setFavorites] = useState<string[]>(() => loadSavedList<string>('orgChartFavorites'));
  const [recentContacts, setRecentContacts] = useState<string[]>(() => loadSavedList<string>('orgChartRecents'));

  // Accordion state
  const [isAddOpen, setIsAddOpen] = useState(false);
  const [isImportOpen, setIsImportOpen] = useState(false);
  const [isDirectoryOpen, setIsDirectoryOpen] = useState(true);

  // Outlook blocks window.confirm, so removals are confirmed inline
  const [pendingRemoval, setPendingRemoval] = useState<Person['id'] | 'all' | null>(null);

  // Toast state
  const [toast, setToast] = useState<{ show: boolean, message: string, type: 'success' | 'error' | 'info' }>({
    show: false, message: '', type: 'info'
  });

  useEffect(() => {
    localStorage.setItem('orgChartData', JSON.stringify(orgData));
  }, [orgData]);

  useEffect(() => {
    localStorage.setItem('orgChartFavorites', JSON.stringify(favorites));
  }, [favorites]);

  useEffect(() => {
    localStorage.setItem('orgChartRecents', JSON.stringify(recentContacts));
  }, [recentContacts]);

  const showToast = (message: string, type: 'success' | 'error' | 'info' = 'info') => {
    setToast({ show: true, message, type });
    setTimeout(() => {
      setToast(prev => ({ ...prev, show: false }));
    }, 3000);
  };

  const addPerson = () => {
    const trimmedEmail = email.trim();

    if (!name.trim() || !trimmedEmail) {
      showToast('Please enter at least name and email', 'error');
      return;
    }

    if (!isValidEmail(trimmedEmail)) {
      showToast('Please enter a valid email address', 'error');
      return;
    }

    if (orgData.some(p => sameEmail(p.email, trimmedEmail))) {
      showToast('Person with this email already exists', 'error');
      return;
    }

    const newPerson: Person = {
      id: Date.now().toString(),
      name: name.trim(), title, email: trimmedEmail, department, reportsTo
    };

    setOrgData([...orgData, newPerson]);
    setName(''); setTitle(''); setEmail(''); setDepartment(''); setReportsTo('');
    showToast('Person added successfully', 'success');
  };

  // Adds rows that have a name and a real email; returns how many were added and how many were invalid
  const mergePeople = (rows: Omit<Person, 'id'>[]) => {
    const newOrgData = [...orgData];
    let imported = 0;
    let invalid = 0;

    rows.forEach(row => {
      if (!row.name || !isValidEmail(row.email)) {
        invalid++;
      } else if (!newOrgData.some(p => sameEmail(p.email, row.email))) {
        newOrgData.push({
          ...row,
          id: Date.now().toString() + Math.random(),
          department: row.department || 'Other'
        });
        imported++;
      }
    });

    // A manager column may hold a name instead of an email
    const resolved = newOrgData.map(p => {
      if (!p.reportsTo || isValidEmail(p.reportsTo)) return p;
      const manager = newOrgData.find(m => m.name.toLowerCase() === p.reportsTo?.toLowerCase());
      return manager ? { ...p, reportsTo: manager.email } : p;
    });

    if (imported > 0) setOrgData(resolved);
    return { imported, invalid };
  };

  const importSummary = (imported: number, invalid: number) =>
    `Imported ${imported} people` + (invalid > 0 ? `, skipped ${invalid} without a valid name and email` : '');

  const importData = () => {
    if (!importText.trim()) {
      showToast('Please enter data to import', 'error');
      return;
    }

    const rows = importText.trim().split('\n').map(line => {
      const parts = line.split(',').map(part => part.trim());
      return {
        name: parts[0] || '',
        title: parts[1] || '',
        email: parts[2] || '',
        department: parts[3] || '',
        reportsTo: parts[4] || ''
      };
    });

    const { imported, invalid } = mergePeople(rows);

    if (imported > 0) {
      setImportText('');
      showToast(importSummary(imported, invalid), 'success');
    } else {
      showToast('No valid new data to import. Each line needs a name and an email.', 'error');
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const file = e.target.files[0];
      const fileName = file.name.toLowerCase();
      
      if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.csv')) {
        const reader = new FileReader();
        reader.onload = async (event) => {
          try {
            const XLSX = await import('xlsx');
            const data = new Uint8Array(event.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON array
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' }) as any[];
            
            const rows = jsonData.map(row => {
              // Standardize column names (allow lowercase/spaces in headers)
              const getCol = (names: string[]) => {
                const key = Object.keys(row).find(k => names.some(n => k.toLowerCase().replace(/\s+/g, '') === n.toLowerCase()));
                return key ? String(row[key]).trim() : '';
              };

              return {
                name: getCol(['name', 'fullname', 'employee']),
                title: getCol(['title', 'jobtitle', 'position', 'role']),
                email: getCol(['email', 'emailaddress']),
                department: getCol(['department', 'dept', 'team']),
                reportsTo: getCol(['reportstoemail', 'manager', 'manageremail', 'reportsto'])
              };
            });

            const { imported, invalid } = mergePeople(rows);

            if (imported > 0) {
              showToast(`${importSummary(imported, invalid)} from ${file.name}`, 'success');
            } else {
              showToast('No new valid entries found in the file.', 'error');
            }
          } catch (err) {
            console.error(err);
            showToast('Error parsing file. Expected columns: Name, Title, Email, Department, ReportsToEmail.', 'error');
          }
        };
        reader.readAsArrayBuffer(file);
      } else {
         showToast(`${file.name} is not supported. Please upload an Excel or CSV file.`, 'error');
      }
      e.target.value = '';
    }
  };

  const removePersonFromOrg = (id: string | number) => {
    setOrgData(orgData.filter(p => p.id !== id));
    setSelectedTo(selectedTo.filter(p => p.id !== id));
    setSelectedCc(selectedCc.filter(p => p.id !== id));
    setPendingRemoval(null);
  };

  const clearAllData = () => {
    setOrgData([]);
    setSelectedTo([]);
    setSelectedCc([]);
    setPendingRemoval(null);
    showToast('All data cleared', 'info');
  };

  const confirmButtons = (label: string, onConfirm: () => void) => (
    <div style={{ display: 'flex', gap: '6px', flexShrink: 0 }} onClick={e => e.stopPropagation()}>
      <button className="btn btn-danger" style={{ padding: '6px 12px', fontSize: '12px' }} onClick={onConfirm}>{label}</button>
      <button className="btn btn-outline" style={{ padding: '6px 12px', fontSize: '12px' }} onClick={() => setPendingRemoval(null)}>Cancel</button>
    </div>
  );

  const getAllAncestors = (person: Person, allData: Person[]): Person[] => {
    const ancestors: Person[] = [];
    const seen = new Set([person.id]);
    let manager = allData.find(p => sameEmail(p.email, person.reportsTo));
    // Imported data can contain reporting loops, so stop at the first repeat
    while (manager && !seen.has(manager.id)) {
      seen.add(manager.id);
      ancestors.push(manager);
      const reportsTo = manager.reportsTo;
      manager = allData.find(p => sameEmail(p.email, reportsTo));
    }
    return ancestors;
  };

  // Removes people from To and Cc, along with managers who were only Cc'd because of them
  const removeFromSelection = (people: Person[]) => {
    const ids = new Set(people.map(p => p.id));
    const remainingTo = selectedTo.filter(p => !ids.has(p.id));
    const ancestorIds = (list: Person[]) => new Set(list.flatMap(p => getAllAncestors(p, orgData)).map(a => a.id));
    const noLongerNeeded = ancestorIds(people);
    const stillNeeded = ancestorIds(remainingTo);

    setSelectedTo(remainingTo);
    setSelectedCc(selectedCc.filter(c => !ids.has(c.id) && (!noLongerNeeded.has(c.id) || stillNeeded.has(c.id))));
  };

  const toggleFavorite = (e: React.MouseEvent, email: string) => {
    e.stopPropagation();
    if (favorites.includes(email)) {
      setFavorites(favorites.filter(f => f !== email));
    } else {
      setFavorites([...favorites, email]);
    }
  };

  const trackRecentAccess = (email: string) => {
    setRecentContacts(prev => {
      const filtered = prev.filter(e => e !== email);
      return [email, ...filtered].slice(0, 8); // Keep up to 8 recent contacts
    });
  };

  const handlePersonClick = (person: Person, listType?: 'to' | 'cc') => {
    trackRecentAccess(person.email);
    if (listType === 'to') {
      removeFromSelection([person]);
    } else if (listType === 'cc') {
      setSelectedCc(selectedCc.filter(p => p.id !== person.id));
    } else {
      const inTo = selectedTo.some(p => p.id === person.id);
      const inCc = selectedCc.some(p => p.id === person.id);
      
      if (inTo || inCc) {
        removeFromSelection([person]);
      } else {
        setSelectedTo([...selectedTo, person]);
        if (ccAncestors) {
          const ancestors = getAllAncestors(person, orgData);
          const newCc = ancestors.filter(a => 
            !selectedCc.some(c => c.id === a.id) && 
            !selectedTo.some(t => t.id === a.id)
          );
          setSelectedCc([...selectedCc, ...newCc]);
        }
      }
    }
  };

  const addToEmail = async () => {
    if (selectedTo.length === 0 && selectedCc.length === 0) {
      showToast('No recipients selected', 'error');
      return;
    }

    const toRecipients = selectedTo.map(p => ({
      displayName: p.name,
      emailAddress: p.email
    }));
    
    const ccRecipients = selectedCc.map(p => ({
      displayName: p.name,
      emailAddress: p.email
    }));

    if (window.Office && window.Office.context && window.Office.context.mailbox) {
      const item = window.Office.context.mailbox.item;

      // Recipients can only be added while composing; in read mode item.to is a plain array
      if (!item?.to?.addAsync) {
        showToast('Open a new email or reply to add recipients', 'info');
        return;
      }

      const addRecipients = (field: any, recipients: typeof toRecipients) =>
        new Promise<boolean>(resolve => {
          if (recipients.length === 0) {
            resolve(true);
            return;
          }
          field.addAsync(recipients, (result: any) => {
            resolve(result.status === window.Office.AsyncResultStatus.Succeeded);
          });
        });

      const [addedTo, addedCc] = await Promise.all([
        addRecipients(item.to, toRecipients),
        addRecipients(item.cc, ccRecipients)
      ]);

      // Keep only what failed selected, so a retry doesn't add the rest twice
      if (addedTo) setSelectedTo([]);
      if (addedCc) setSelectedCc([]);

      if (addedTo && addedCc) {
        showToast(`Added ${toRecipients.length} TO and ${ccRecipients.length} CC recipients`, 'success');
      } else {
        showToast(`Error adding ${addedTo ? 'CC' : addedCc ? 'TO' : 'TO and CC'} recipients`, 'error');
      }
    } else {
      // Not running in Outlook
      showToast(`(Simulation) Added ${toRecipients.length} TO and ${ccRecipients.length} CC recipients`, 'success');
      setSelectedTo([]);
      setSelectedCc([]);
    }
  };

  const selectDepartment = (_dept: string, peopleInDept: Person[]) => {
    const allSelected = peopleInDept.every(p => selectedTo.some(sp => sp.id === p.id));
    if (allSelected) {
      removeFromSelection(peopleInDept);
    } else {
      const currentlySelectedIds = new Set(selectedTo.map(p => p.id));
      const newPeople = peopleInDept.filter(p => !currentlySelectedIds.has(p.id));
      setSelectedTo([...selectedTo, ...newPeople]);

      // Anyone moving to To shouldn't stay in Cc as well
      const newIds = new Set(newPeople.map(p => p.id));
      setSelectedCc(prev => prev.filter(c => !newIds.has(c.id)));

      if (ccAncestors) {
        const allAncestors = newPeople.flatMap(p => getAllAncestors(p, orgData));
        const uniqueAncestors = Array.from(new Map(allAncestors.map(a => [a.id, a])).values());
        
        const newCc = uniqueAncestors.filter(a => 
          !selectedCc.some(c => c.id === a.id) && 
          !selectedTo.some(t => t.id === a.id) &&
          !newPeople.some(np => np.id === a.id)
        );
        setSelectedCc(prev => [...prev, ...newCc]);
      }
    }
  };

  // Group by department for the Use tab
  const filteredData = orgData.filter(p =>
    p.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
    p.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
    p.email.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const departments = filteredData.reduce((acc, person) => {
    const dept = person.department || 'Other';
    if (!acc[dept]) acc[dept] = [];
    acc[dept].push(person);
    return acc;
  }, {} as Record<string, Person[]>);

  // Shared across the tree render so reporting loops can't recurse forever or hide people
  const renderedInTree = new Set<Person['id']>();

  return (
    <>
      <div className={`toast ${toast.show ? 'show' : ''} ${toast.type}`}>
        {toast.type === 'success' && <CheckCircle2 size={18} />}
        {toast.type === 'error' && <XCircle size={18} />}
        {toast.type === 'info' && <Users size={18} />}
        <span>{toast.message}</span>
      </div>

      <div className="app-header">
        <h1>OrgChartCC</h1>
        <h2>Interactive Org Chart</h2>
        <p>Looping in the right people - every time!</p>
      </div>

      <div className="tabs-container">
        <button
          className={`tab-btn ${activeTab === 'create' ? 'active' : ''}`}
          onClick={() => setActiveTab('create')}
        >
          <Building2 size={16} /> Add to your OrgChart
        </button>
        <button
          className={`tab-btn ${activeTab === 'use' ? 'active' : ''}`}
          onClick={() => setActiveTab('use')}
        >
          <Users size={16} /> Select Recipients
        </button>
      </div>

      {activeTab === 'create' && (
        <>
          <div className="section-card">
            <div className="section-title collapsible-title" onClick={() => setIsAddOpen(!isAddOpen)}>
              <span style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                {isAddOpen ? <ChevronDown size={18} /> : <ChevronRight size={18} />} Add Individual
              </span>
            </div>
            {isAddOpen && (
              <div className="collapsible-content">
                <div className="form-group">
                  <label className="form-label">Full Name</label>
                  <input type="text" className="form-input" value={name} onChange={e => setName(e.target.value)} placeholder="John Doe" />
                </div>
                <div className="form-group">
                  <label className="form-label">Job Title</label>
                  <input type="text" className="form-input" value={title} onChange={e => setTitle(e.target.value)} placeholder="Senior Developer" />
                </div>
                <div className="form-group">
                  <label className="form-label">Email Address</label>
                  <input type="email" className="form-input" value={email} onChange={e => setEmail(e.target.value)} placeholder="john@company.com" />
                </div>
                <div className="form-group">
                  <label className="form-label">Department</label>
                  <select className="form-select" value={department} onChange={e => setDepartment(e.target.value)}>
                    <option value="">Select department</option>
                    <option value="Executive">Executive</option>
                    <option value="Engineering">Engineering</option>
                    <option value="Sales">Sales</option>
                    <option value="Marketing">Marketing</option>
                    <option value="Human Resources">Human Resources</option>
                    <option value="Finance">Finance</option>
                    <option value="Operations">Operations</option>
                  </select>
                </div>
                <div className="form-group">
                  <label className="form-label">Reports To</label>
                  <select className="form-select" value={reportsTo} onChange={e => setReportsTo(e.target.value)}>
                    <option value="">None</option>
                    {orgData.map(p => (
                      <option key={p.id} value={p.email}>{p.name} ({p.title})</option>
                    ))}
                  </select>
                </div>
                <button className="btn btn-primary btn-full" onClick={addPerson}>
                  <UserPlus size={16} /> Add to Directory
                </button>
              </div>
            )}
          </div>

          <div className="section-card">
            <div className="section-title collapsible-title" onClick={() => setIsImportOpen(!isImportOpen)}>
              <span style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                {isImportOpen ? <ChevronDown size={18} /> : <ChevronRight size={18} />} Import your Org Chart
              </span>
            </div>
            {isImportOpen && (
              <div className="collapsible-content">
                <div style={{ marginBottom: '12px', fontSize: '13px', fontWeight: '500' }}>Option 1: Paste text data (CSV format)</div>
                <div className="form-group">
                  <textarea
                    className="form-textarea"
                    rows={6}
                    value={importText}
                    onChange={e => setImportText(e.target.value)}
                    placeholder="Name, Title, Email, Department, ReportsToEmail"
                  />
                </div>
                <button className="btn btn-outline btn-full" onClick={importData} style={{ marginBottom: '24px' }}>
                  <Upload size={16} /> Import Text Data
                </button>

                <div style={{ marginBottom: '12px', fontSize: '13px', fontWeight: '500' }}>Option 2: Upload file (Excel or CSV)</div>
                <div className="form-group">
                  <input 
                    type="file" 
                    accept=".xlsx, .xls, .csv"
                    className="form-input" 
                    style={{ padding: '10px' }}
                    onChange={handleFileUpload} 
                  />
                  <p style={{ fontSize: '12px', color: 'var(--text-muted)', marginTop: '6px' }}>Columns: Name, Title, Email, Department, ReportsToEmail. Name and Email are required.</p>
                </div>
              </div>
            )}
          </div>

          <div className="section-card">
            <div className="section-title collapsible-title" onClick={() => setIsDirectoryOpen(!isDirectoryOpen)}>
              <span style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                {isDirectoryOpen ? <ChevronDown size={18} /> : <ChevronRight size={18} />} Directory ({orgData.length})
              </span>
              {orgData.length > 0 && (pendingRemoval === 'all' ? confirmButtons('Clear all', clearAllData) : (
                <button className="btn btn-danger" style={{ padding: '6px 12px', fontSize: '12px' }} onClick={(e) => { e.stopPropagation(); setPendingRemoval('all'); }}>
                  Clear
                </button>
              ))}
            </div>
            
            {isDirectoryOpen && (
              <div className="collapsible-content">
                <div className="person-list">
                  {orgData.length === 0 ? (
                    <p style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '20px' }}>No people in directory</p>
                  ) : (
                    orgData.map(person => (
                      <div key={person.id} className="person-item">
                        <div className="person-info">
                          <span className="person-name">{person.name}</span>
                          <span className="person-title">{person.title} {person.department ? `· ${person.department}` : ''}</span>
                          <span className="person-email">{person.email}</span>
                        </div>
                        {pendingRemoval === person.id ? confirmButtons('Delete', () => removePersonFromOrg(person.id)) : (
                          <button className="btn btn-outline" style={{ padding: '6px', color: 'var(--danger)', border: 'none' }} onClick={() => setPendingRemoval(person.id)}>
                            <Trash2 size={16} />
                          </button>
                        )}
                      </div>
                    ))
                  )}
                </div>
              </div>
            )}
          </div>
        </>
      )}

      {activeTab === 'use' && (
        <>
          <div className="view-mode-toggle">
            <button
              className={`view-btn ${viewMode === 'department' ? 'active' : ''}`}
              onClick={() => setViewMode('department')}
            >
              View by department
            </button>
            <button
              className={`view-btn ${viewMode === 'orgchart' ? 'active' : ''}`}
              onClick={() => setViewMode('orgchart')}
            >
              View whole org chart
            </button>
          </div>
          
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '16px', padding: '0 4px' }}>
            <input 
              type="checkbox" 
              id="ccAncestors" 
              checked={ccAncestors} 
              onChange={e => setCcAncestors(e.target.checked)} 
              style={{ width: '16px', height: '16px', cursor: 'pointer' }}
            />
            <label htmlFor="ccAncestors" style={{ fontSize: '14px', cursor: 'pointer', userSelect: 'none' }}>
              CC everyone above the position
            </label>
          </div>

          <div className="form-group">
            <input
              type="text"
              className="form-input"
              placeholder="Search by name, title, or email..."
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
              style={{ borderRadius: '24px', padding: '12px 20px', boxShadow: 'var(--shadow-sm)' }}
            />
          </div>

          {/* Quick Access Section */}
          {(favorites.length > 0 || recentContacts.length > 0) && (
            <div className="quick-access-section">
              {favorites.length > 0 && (
                <div className="quick-access-group">
                  <div className="quick-access-title"><Star size={14} className="icon-gold" /> Favorites</div>
                  <div className="quick-access-list">
                    {favorites.map(email => {
                      const person = orgData.find(p => p.email === email);
                      if (!person) return null;
                      const isSelected = selectedTo.some(p => p.id === person.id) || selectedCc.some(p => p.id === person.id);
                      return (
                        <div key={`fav-${person.id}`} className={`quick-person-card ${isSelected ? 'selected' : ''}`} onClick={() => handlePersonClick(person)}>
                          <div className="quick-person-avatar">{person.name.charAt(0)}</div>
                          <div className="quick-person-info">
                            <span className="name">{person.name}</span>
                            <span className="title">{person.title}</span>
                          </div>
                          <button className="favorite-btn active" onClick={(e) => toggleFavorite(e, person.email)}>
                            <Star size={14} fill="currentColor" />
                          </button>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
              
              {recentContacts.length > 0 && (
                <div className="quick-access-group">
                  <div className="quick-access-title"><History size={14} /> Recently Used</div>
                  <div className="quick-access-list">
                    {recentContacts.map(email => {
                      const person = orgData.find(p => p.email === email);
                      if (!person) return null;
                      const isSelected = selectedTo.some(p => p.id === person.id) || selectedCc.some(p => p.id === person.id);
                      const isFav = favorites.includes(person.email);
                      return (
                        <div key={`rec-${person.id}`} className={`quick-person-card ${isSelected ? 'selected' : ''}`} onClick={() => handlePersonClick(person)}>
                          <div className="quick-person-avatar">{person.name.charAt(0)}</div>
                          <div className="quick-person-info">
                            <span className="name">{person.name}</span>
                            <span className="title">{person.title}</span>
                          </div>
                          <button className={`favorite-btn ${isFav ? 'active' : ''}`} onClick={(e) => toggleFavorite(e, person.email)}>
                            <Star size={14} fill={isFav ? "currentColor" : "none"} />
                          </button>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          )}

          {(selectedTo.length > 0 || selectedCc.length > 0) && (
            <div className="selected-tags-container">
              {selectedTo.map(person => (
                <div key={`sel-to-${person.id}`} className="selected-tag to-tag">
                  <span style={{ fontWeight: 'bold', marginRight: '4px' }}>To:</span> {person.name}
                  <button onClick={() => handlePersonClick(person, 'to')}><XCircle size={14} /></button>
                </div>
              ))}
              {selectedCc.map(person => (
                <div key={`sel-cc-${person.id}`} className="selected-tag cc-tag">
                  <span style={{ fontWeight: 'bold', marginRight: '4px' }}>Cc:</span> {person.name}
                  <button onClick={() => handlePersonClick(person, 'cc')}><XCircle size={14} /></button>
                </div>
              ))}
            </div>
          )}

          <div style={{ display: 'flex', gap: '12px', marginBottom: '24px' }}>
            <button className="btn btn-primary btn-full" onClick={addToEmail} disabled={selectedTo.length === 0 && selectedCc.length === 0}>
              <MailPlus size={16} /> Add to Email ({selectedTo.length + selectedCc.length})
            </button>
            {(selectedTo.length > 0 || selectedCc.length > 0) && (
              <button className="btn btn-outline" onClick={() => { setSelectedTo([]); setSelectedCc([]); }}>Clear</button>
            )}
          </div>

          <div className="org-chart-container-scroll">
            <div className="org-chart-container">
              {Object.keys(departments).length === 0 ? (
                <div className="section-card">
                  <p style={{ textAlign: 'center', color: 'var(--text-muted)' }}>
                    {orgData.length === 0 ? 'No directory data. Please add people first.' : 'No results found.'}
                  </p>
                </div>
              ) : viewMode === 'department' ? (
                Object.keys(departments).sort().map(dept => (
                  <div key={dept}>
                    <div className="dept-header clickable-header" onClick={() => selectDepartment(dept, departments[dept])}>
                      {dept} <span className="dept-select-hint">(Click to select all)</span>
                    </div>
                    <div className="dept-grid">
                      {departments[dept].map(person => {
                        const isSelected = selectedTo.some(p => p.id === person.id) || selectedCc.some(p => p.id === person.id);
                        return (
                          <div
                            key={person.id}
                            className={`user-card ${isSelected ? 'selected' : ''}`}
                            onClick={() => handlePersonClick(person)}
                          >
                            <Check className="check-icon" size={18} />
                            <button className={`card-fav-btn ${favorites.includes(person.email) ? 'active' : ''}`} onClick={(e) => toggleFavorite(e, person.email)}>
                              <Star size={14} fill={favorites.includes(person.email) ? "currentColor" : "none"} />
                            </button>
                            <div className="person-info">
                              <span className="person-name">{person.name}</span>
                              <span className="person-title">{person.title}</span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ))
              ) : (
                <div className="tree-org-chart">
                  <div className="tree-node root-node">
                    <div className="tree-card root-card">Company</div>
                    <div className="tree-children">
                      {[
                        ...filteredData.filter(p => !filteredData.some(other => sameEmail(other.email, p.reportsTo))),
                        ...filteredData
                      ].map(root => {
                        const renderTree = (person: Person) => {
                          renderedInTree.add(person.id);
                          const children = filteredData.filter(p => sameEmail(person.email, p.reportsTo) && !renderedInTree.has(p.id));
                          children.forEach(child => renderedInTree.add(child.id));
                          const isSelected = selectedTo.some(p => p.id === person.id) || selectedCc.some(p => p.id === person.id);
                          return (
                            <div className="tree-node person-node" key={person.id}>
                              <div
                                className={`tree-card person-card ${isSelected ? 'selected' : ''}`}
                                onClick={() => handlePersonClick(person)}
                              >
                                <Check className="check-icon" size={14} />
                                <button className={`card-fav-btn ${favorites.includes(person.email) ? 'active' : ''}`} onClick={(e) => toggleFavorite(e, person.email)}>
                                  <Star size={14} fill={favorites.includes(person.email) ? "currentColor" : "none"} />
                                </button>
                                <span className="person-name">{person.name}</span>
                                <span className="person-title">{person.title}</span>
                              </div>
                              {children.length > 0 && (
                                <div className="tree-children">
                                  {children.map(child => renderTree(child))}
                                </div>
                              )}
                            </div>
                          );
                        };
                        return renderedInTree.has(root.id) ? null : renderTree(root);
                      })}
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        </>
      )}
    </>
  );
}
