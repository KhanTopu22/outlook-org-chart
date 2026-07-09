import { useState, useEffect } from 'react';
import {
  Users, UserPlus, Upload, Trash2, MailPlus,
  CheckCircle2, XCircle, Check, Building2,
  ChevronDown, ChevronRight
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

export default function App() {
  const [activeTab, setActiveTab] = useState<TabType>('use');
  const [viewMode, setViewMode] = useState<ViewMode>('department');
  const [orgData, setOrgData] = useState<Person[]>([]);
  const [selectedPeople, setSelectedPeople] = useState<Person[]>([]);

  // Form state
  const [name, setName] = useState('');
  const [title, setTitle] = useState('');
  const [email, setEmail] = useState('');
  const [department, setDepartment] = useState('');
  const [reportsTo, setReportsTo] = useState('');
  const [importText, setImportText] = useState('');
  const [searchTerm, setSearchTerm] = useState('');

  // Accordion state
  const [isAddOpen, setIsAddOpen] = useState(false);
  const [isImportOpen, setIsImportOpen] = useState(false);
  const [isDirectoryOpen, setIsDirectoryOpen] = useState(true);

  // Toast state
  const [toast, setToast] = useState<{ show: boolean, message: string, type: 'success' | 'error' | 'info' }>({
    show: false, message: '', type: 'info'
  });

  useEffect(() => {
    const saved = localStorage.getItem('orgChartData');
    if (saved) {
      try {
        setOrgData(JSON.parse(saved));
      } catch (e) {
        console.error("Failed to parse saved org chart data");
      }
    }
  }, []);

  const showToast = (message: string, type: 'success' | 'error' | 'info' = 'info') => {
    setToast({ show: true, message, type });
    setTimeout(() => {
      setToast(prev => ({ ...prev, show: false }));
    }, 3000);
  };

  const saveOrgChart = () => {
    localStorage.setItem('orgChartData', JSON.stringify(orgData));
    showToast('Organization chart saved successfully', 'success');
  };

  const addPerson = () => {
    if (!name || !email) {
      showToast('Please enter at least name and email', 'error');
      return;
    }

    if (orgData.some(p => p.email.toLowerCase() === email.toLowerCase())) {
      showToast('Person with this email already exists', 'error');
      return;
    }

    const newPerson: Person = {
      id: Date.now().toString(),
      name, title, email, department, reportsTo
    };
    
    setOrgData([...orgData, newPerson]);
    setName(''); setTitle(''); setEmail(''); setDepartment(''); setReportsTo('');
    showToast('Person added successfully', 'success');
  };

  const importData = () => {
    if (!importText.trim()) {
      showToast('Please enter data to import', 'error');
      return;
    }

    const lines = importText.trim().split('\n');
    let imported = 0;
    const newOrgData = [...orgData];

    lines.forEach((line) => {
      const parts = line.split(',').map(part => part.trim());
      if (parts.length >= 2 && parts[0] && parts[1]) {
        const pEmail = parts[2] || parts[0].toLowerCase().replace(/\s+/g, '.') + '@company.com';

        if (!newOrgData.some(p => p.email.toLowerCase() === pEmail.toLowerCase())) {
          newOrgData.push({
            id: Date.now().toString() + Math.random(),
            name: parts[0],
            title: parts[1],
            email: pEmail,
            department: parts[3] || 'Other',
            reportsTo: parts[4] || ''
          });
          imported++;
        }
      }
    });

    if (imported > 0) {
      setOrgData(newOrgData);
      setImportText('');
      showToast(`Imported ${imported} people`, 'success');
    } else {
      showToast('No valid new data to import', 'error');
    }
  };

  const removePersonFromOrg = (id: string | number) => {
    if (window.confirm('Are you sure you want to remove this person?')) {
      setOrgData(orgData.filter(p => p.id !== id));
      setSelectedPeople(selectedPeople.filter(p => p.id !== id));
    }
  };

  const clearAllData = () => {
    if (window.confirm('Are you sure you want to clear all data?')) {
      setOrgData([]);
      setSelectedPeople([]);
      localStorage.removeItem('orgChartData');
      showToast('All data cleared', 'info');
    }
  };

  const togglePersonSelection = (person: Person) => {
    if (selectedPeople.some(p => p.id === person.id)) {
      setSelectedPeople(selectedPeople.filter(p => p.id !== person.id));
    } else {
      setSelectedPeople([...selectedPeople, person]);
    }
  };

  const getAllDescendants = (person: Person, allData: Person[]): Person[] => {
    const children = allData.filter(p => p.reportsTo === person.email);
    let descendants = [...children];
    children.forEach(child => {
      descendants = [...descendants, ...getAllDescendants(child, allData)];
    });
    return descendants;
  };

  const toggleChainSelection = (person: Person) => {
    const descendants = getAllDescendants(person, orgData);
    const family = [person, ...descendants];
    const familyIds = new Set(family.map(p => p.id));
    const allSelected = family.every(p => selectedPeople.some(sp => sp.id === p.id));
    
    if (allSelected) {
      setSelectedPeople(selectedPeople.filter(p => !familyIds.has(p.id)));
    } else {
      const currentlySelectedIds = new Set(selectedPeople.map(p => p.id));
      const newPeople = family.filter(p => !currentlySelectedIds.has(p.id));
      setSelectedPeople([...selectedPeople, ...newPeople]);
    }
  };

  const addToEmail = () => {
    if (selectedPeople.length === 0) {
      showToast('No recipients selected', 'error');
      return;
    }

    const recipients = selectedPeople.map(p => ({
      displayName: p.name,
      emailAddress: p.email
    }));

    if (window.Office && window.Office.context && window.Office.context.mailbox) {
      window.Office.context.mailbox.item.to.addAsync(recipients, (result: any) => {
        if (result.status === window.Office.AsyncResultStatus.Succeeded) {
          showToast(`Added ${selectedPeople.length} recipients to email`, 'success');
          setSelectedPeople([]);
        } else {
          showToast('Error adding recipients. Please try again.', 'error');
        }
      });
    } else {
      // Not running in Outlook
      showToast(`(Simulation) Added ${selectedPeople.length} recipients to email`, 'success');
      setSelectedPeople([]);
    }
  };

  const selectDepartment = (_dept: string, peopleInDept: Person[]) => {
    const allSelected = peopleInDept.every(p => selectedPeople.some(sp => sp.id === p.id));
    if (allSelected) {
      const idsToRemove = new Set(peopleInDept.map(p => p.id));
      setSelectedPeople(selectedPeople.filter(p => !idsToRemove.has(p.id)));
    } else {
      const currentlySelectedIds = new Set(selectedPeople.map(p => p.id));
      const newPeople = peopleInDept.filter(p => !currentlySelectedIds.has(p.id));
      setSelectedPeople([...selectedPeople, ...newPeople]);
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
                    <option value="HR">Human Resources</option>
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
                <div className="form-group">
                  <textarea
                    className="form-textarea"
                    rows={8}
                    value={importText}
                    onChange={e => setImportText(e.target.value)}
                    placeholder="Alice, CEO, ceo@gameco.com, Executive, &#10;Bob, Office Manager, admin@gameco.com, Administration, ceo@gameco.com&#10;Charlie, Art Director, art@gameco.com, Artwork Creation, ceo@gameco.com&#10;Dave, Dir. of Technology, tech@gameco.com, Programming, ceo@gameco.com&#10;Eve, Dir. of Operations, ops@gameco.com, Production, ceo@gameco.com&#10;Frank, Lead Artist, leadart@gameco.com, Artwork Creation, art@gameco.com&#10;Grace, Gameplay Programmer, dev@gameco.com, Programming, tech@gameco.com"
                  />
                </div>
                <button className="btn btn-outline btn-full" onClick={importData}>
                  <Upload size={16} /> Import Data
                </button>
              </div>
            )}
          </div>

          <div className="section-card">
            <div className="section-title collapsible-title" onClick={() => setIsDirectoryOpen(!isDirectoryOpen)}>
              <span style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                {isDirectoryOpen ? <ChevronDown size={18} /> : <ChevronRight size={18} />} Directory ({orgData.length})
              </span>
              {orgData.length > 0 && (
                <button className="btn btn-danger" style={{ padding: '6px 12px', fontSize: '12px' }} onClick={(e) => { e.stopPropagation(); clearAllData(); }}>
                  Clear
                </button>
              )}
            </div>
            
            {isDirectoryOpen && (
              <div className="collapsible-content">
                <div className="person-list" style={{ marginBottom: '16px' }}>
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
                        <button className="btn btn-outline" style={{ padding: '6px', color: 'var(--danger)', border: 'none' }} onClick={() => removePersonFromOrg(person.id)}>
                          <Trash2 size={16} />
                        </button>
                      </div>
                    ))
                  )}
                </div>
                
                <button className="btn btn-success btn-full" onClick={saveOrgChart} disabled={orgData.length === 0}>
                  <CheckCircle2 size={16} /> Save Directory
                </button>
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

          {selectedPeople.length > 0 && (
            <div className="selected-tags-container">
              {selectedPeople.map(person => (
                <div key={`sel-${person.id}`} className="selected-tag">
                  {person.name}
                  <button onClick={() => togglePersonSelection(person)}><XCircle size={14} /></button>
                </div>
              ))}
            </div>
          )}

          <div style={{ display: 'flex', gap: '12px', marginBottom: '24px' }}>
            <button className="btn btn-primary btn-full" onClick={addToEmail} disabled={selectedPeople.length === 0}>
              <MailPlus size={16} /> Add to Email ({selectedPeople.length})
            </button>
            {selectedPeople.length > 0 && (
              <button className="btn btn-outline" onClick={() => setSelectedPeople([])}>Clear</button>
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
                        const isSelected = selectedPeople.some(p => p.id === person.id);
                        return (
                          <div
                            key={person.id}
                            className={`user-card ${isSelected ? 'selected' : ''}`}
                            onClick={() => togglePersonSelection(person)}
                          >
                            <Check className="check-icon" size={18} />
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
                      {orgData.filter(p => !p.reportsTo || !orgData.some(other => other.email === p.reportsTo)).map(root => {
                        const renderTree = (person: Person) => {
                          const children = orgData.filter(p => p.reportsTo === person.email);
                          const isSelected = selectedPeople.some(p => p.id === person.id);
                          return (
                            <div className="tree-node person-node" key={person.id}>
                              <div
                                className={`tree-card person-card ${isSelected ? 'selected' : ''}`}
                                onClick={() => toggleChainSelection(person)}
                              >
                                <Check className="check-icon" size={14} />
                                <span className="person-name">{person.name}</span>
                                <span className="person-title">{person.title}</span>
                                {children.length > 0 && <span className="dept-select-hint">(Select team)</span>}
                              </div>
                              {children.length > 0 && (
                                <div className="tree-children">
                                  {children.map(child => renderTree(child))}
                                </div>
                              )}
                            </div>
                          );
                        };
                        return renderTree(root);
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
