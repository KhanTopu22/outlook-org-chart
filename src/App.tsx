import { useState, useEffect } from 'react';
import { 
  Users, UserPlus, Upload, Trash2, MailPlus, 
  CheckCircle2, XCircle, Check, Building2
} from 'lucide-react';
import './index.css';

// Types
interface Person {
  id: string | number;
  name: string;
  title: string;
  email: string;
  department: string;
}

type TabType = 'create' | 'use';

export default function App() {
  const [activeTab, setActiveTab] = useState<TabType>('create');
  const [orgData, setOrgData] = useState<Person[]>([]);
  const [selectedPeople, setSelectedPeople] = useState<Person[]>([]);
  
  // Form state
  const [name, setName] = useState('');
  const [title, setTitle] = useState('');
  const [email, setEmail] = useState('');
  const [department, setDepartment] = useState('');
  const [importText, setImportText] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  
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

  const loadTemplate = (type: 'startup' | 'corporate') => {
    if (orgData.length > 0 && !window.confirm('This will replace your current org chart. Continue?')) {
      return;
    }

    let templateData: Person[] = [];
    if (type === 'startup') {
      templateData = [
        {id: '1', name: 'Alex Johnson', title: 'CEO', email: 'alex@company.com', department: 'Executive'},
        {id: '2', name: 'Sarah Chen', title: 'CTO', email: 'sarah@company.com', department: 'Engineering'},
        {id: '3', name: 'Mike Davis', title: 'VP Sales', email: 'mike@company.com', department: 'Sales'},
        {id: '4', name: 'Lisa Wang', title: 'Lead Developer', email: 'lisa@company.com', department: 'Engineering'},
        {id: '5', name: 'Tom Brown', title: 'Sales Manager', email: 'tom@company.com', department: 'Sales'}
      ];
    } else {
      templateData = [
        {id: '1', name: 'Robert Smith', title: 'CEO', email: 'robert@company.com', department: 'Executive'},
        {id: '2', name: 'David Lee', title: 'CTO', email: 'david@company.com', department: 'Engineering'},
        {id: '3', name: 'Maria Garcia', title: 'VP Marketing', email: 'maria@company.com', department: 'Marketing'},
        {id: '4', name: 'Anna Kim', title: 'HR Director', email: 'anna@company.com', department: 'HR'},
        {id: '5', name: 'Chris Martinez', title: 'Operations', email: 'chris@company.com', department: 'Operations'},
      ];
    }
    setOrgData(templateData);
    showToast(`Loaded ${type} template`, 'success');
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
      name, title, email, department
    };
    
    setOrgData([...orgData, newPerson]);
    setName(''); setTitle(''); setEmail(''); setDepartment('');
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
            department: parts[3] || 'Other'
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
        <h2>Interactive Org Chart</h2>
        <p>Select recipients seamlessly from your organizational structure</p>
      </div>

      <div className="tabs-container">
        <button 
          className={`tab-btn ${activeTab === 'create' ? 'active' : ''}`}
          onClick={() => setActiveTab('create')}
        >
          <Building2 size={16} /> Manage Organization
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
          <div className="section-card" style={{ background: 'linear-gradient(to right, rgba(0,120,212,0.05), transparent)' }}>
            <div className="section-title">
              Quick Start Templates
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              <button className="btn btn-outline" onClick={() => loadTemplate('startup')}>Startup (5 people)</button>
              <button className="btn btn-outline" onClick={() => loadTemplate('corporate')}>Corporate (15 people)</button>
            </div>
          </div>

          <div className="section-card">
            <div className="section-title">Add Individual</div>
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
            <button className="btn btn-primary btn-full" onClick={addPerson}>
              <UserPlus size={16} /> Add to Directory
            </button>
          </div>

          <div className="section-card">
            <div className="section-title">Bulk Import (CSV)</div>
            <div className="form-group">
              <label className="form-label">Format: Name, Title, Email, Department</label>
              <textarea 
                className="form-textarea" 
                rows={3} 
                value={importText} 
                onChange={e => setImportText(e.target.value)}
                placeholder="Jane Doe, VP Sales, jane@company.com, Sales" 
              />
            </div>
            <button className="btn btn-outline btn-full" onClick={importData}>
              <Upload size={16} /> Import Data
            </button>
          </div>

          <div className="section-card">
            <div className="section-title" style={{ justifyContent: 'space-between' }}>
              <span>Directory ({orgData.length})</span>
              {orgData.length > 0 && (
                <button className="btn btn-danger" style={{ padding: '6px 12px', fontSize: '12px' }} onClick={clearAllData}>
                  Clear
                </button>
              )}
            </div>
            
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
        </>
      )}

      {activeTab === 'use' && (
        <>
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

          <div className="org-chart-container">
            {Object.keys(departments).length === 0 ? (
              <div className="section-card">
                <p style={{ textAlign: 'center', color: 'var(--text-muted)' }}>
                  {orgData.length === 0 ? 'No directory data. Please add people first.' : 'No results found.'}
                </p>
              </div>
            ) : (
              Object.keys(departments).sort().map(dept => (
                <div key={dept}>
                  <div className="dept-header">{dept}</div>
                  <div className="dept-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '12px' }}>
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
                            <span className="person-email">{person.email}</span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))
            )}
          </div>
        </>
      )}
    </>
  );
}
