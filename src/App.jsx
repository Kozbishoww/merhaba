import React, { useState, useEffect, useMemo } from 'react';
import { 
  Folder, FolderOpen, FileText, Plus, Edit, Trash2, 
  Search, Printer, X, ChevronRight, ChevronDown, 
  User, Settings, Download, Lock, Eye, EyeOff, Shield, Key,
  RefreshCw, CheckCircle, BookOpen, BarChart2, ArrowRightLeft, FileSpreadsheet
} from 'lucide-react';
import * as XLSX from 'xlsx';
import './index.css';

// --- INITIAL MOCK USERS & FIRMS ---
const initialFirms = [
  { id: '001', name: 'OZN TİCARET VE SANAYİ A.Ş.', period: '01 (2026)', taxOffice: 'Büyük Mükellefler V.D.', taxNo: '6480912345', createdAt: '16.08.2026' }
];

const initialUsers = [
  { username: 'admin', password: '123', role: 'admin', authorizedFirms: [] },
  { username: 'veli', password: '123', role: 'user', authorizedFirms: ['001'] }
];

// --- TEK DÜZEN HESAP PLANI VARSAYILAN ---
const defaultHesapPlani = [
  { code: '100', desc: 'KASA HESABI', type: 'Aktif', borc: 0, alacak: 0 },
  { code: '101', desc: 'ALINAN ÇEKLER', type: 'Aktif', borc: 0, alacak: 0 },
  { code: '102', desc: 'BANKALAR', type: 'Aktif', borc: 0, alacak: 0 },
  { code: '103', desc: 'VERİLEN ÇEKLER VE ÖDEME EMİRLERİ (-)', type: 'Aktif', borc: 0, alacak: 0 },
  { code: '120', desc: 'ALICILAR', type: 'Aktif', borc: 0, alacak: 0 },
  { code: '121', desc: 'ALACAK SENETLERİ', type: 'Aktif', borc: 0, alacak: 0 },
  { code: '153', desc: 'TİCARİ MALLAR', type: 'Aktif', borc: 0, alacak: 0 },
  { code: '191', desc: 'İNDİRİLECEK KDV', type: 'Aktif', borc: 0, alacak: 0 },
  { code: '320', desc: 'SATICILAR', type: 'Pasif', borc: 0, alacak: 0 },
  { code: '321', desc: 'BORÇ SENETLERİ', type: 'Pasif', borc: 0, alacak: 0 },
  { code: '391', desc: 'HESAPLANAN KDV', type: 'Pasif', borc: 0, alacak: 0 },
  { code: '500', desc: 'SERMAYE', type: 'Pasif', borc: 0, alacak: 0 },
  { code: '600', desc: 'YURTİÇİ SATIŞLAR', type: 'Gelir', borc: 0, alacak: 0 },
  { code: '610', desc: 'SATIŞTAN İADELER (-)', type: 'Gider', borc: 0, alacak: 0 },
  { code: '621', desc: 'SATILAN TİCARİ MALLAR MALİYETİ (-)', type: 'Gider', borc: 0, alacak: 0 },
  { code: '770', desc: 'GENEL YÖNETİM GİDERLERİ', type: 'Gider', borc: 0, alacak: 0 }
];

// --- LOGO J-PLUS / J-PLATFORM MENU AĞACI ---
const erpMenuData = [
  { id: 'malzeme_yonetimi', title: 'Malzeme Yönetimi', type: 'folder', isOpen: true, children: [ 
    { id: 'malzemeler', title: 'Malzemeler (Stok Kartları)', type: 'item' },
    { id: 'malzeme_fisleri', title: 'Malzeme Yönetim Fişleri', type: 'item' }
  ] },
  { id: 'satinalma', title: 'Satınalma Yönetimi', type: 'folder', isOpen: true, children: [ 
    { id: 'satinalma_siparisleri', title: 'Satınalma Siparişleri', type: 'item' },
    { id: 'satinalma_faturalari', title: 'Satınalma Faturaları', type: 'item' }
  ] },
  { id: 'satis_dagitim', title: 'Satış ve Dağıtım', type: 'folder', isOpen: true, children: [ 
    { id: 'satis_siparisleri', title: 'Satış Siparişleri', type: 'item' },
    { id: 'satis_faturalari', title: 'Satış Faturaları', type: 'item' }
  ] },
  { id: 'finans', title: 'Finans Yönetimi', type: 'folder', isOpen: true, children: [ 
    { id: 'cari_hesaplar', title: 'Cari Hesaplar', type: 'item' },
    { id: 'cari_hesap_fisleri', title: 'Cari Hesap Fişleri', type: 'item' },
    { id: 'kasa_kartlari', title: 'Kasa Kartları', type: 'item' },
    { id: 'kasa_islemleri', title: 'Kasa İşlemleri (Fişler)', type: 'item' },
    { id: 'banka_kartlari', title: 'Banka ve Hesap Kartları', type: 'item' },
    { id: 'banka_islemleri', title: 'Banka İşlemleri (Havale/EFT)', type: 'item' },
    { id: 'cek_senet', title: 'Çek ve Senet Bordroları', type: 'item' }
  ] },
  { id: 'genel_muhasebe', title: 'Genel Muhasebe', type: 'folder', isOpen: true, children: [ 
    { id: 'hesap_plani', title: 'Muhasebe Hesap Planı', type: 'item' },
    { id: 'muhasebe_fisleri', title: 'Muhasebe Fişleri (Yevmiye)', type: 'item' },
    { id: 'mizan_raporu', title: 'Genel Mizan ve Mali Tablo', type: 'item' }
  ] }
];

const adminMenuData = [
  { id: 'kurulus', title: 'Kuruluş Bilgileri', type: 'folder', isOpen: true, children: [ 
    { id: 'firmalar', title: 'Firmalar ve Dönemler', type: 'item' }
  ] },
  { id: 'yonetim', title: 'Yönetim ve Yetki', type: 'folder', isOpen: true, children: [ 
    { id: 'kullanicilar', title: 'Kullanıcılar ve Firma Atama', type: 'item' }
  ] }
];

const formatMoney = (val) => {
  const num = parseFloat(val) || 0;
  return num.toLocaleString('tr-TR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

// ==========================================
// LOGIN COMPONENT
// ==========================================
const LoginScreen = ({ onLogin, onFirmSelect, pendingUser, firms }) => {
  const [username, setUsername] = useState('veli');
  const [password, setPassword] = useState('123');
  const [showPassword, setShowPassword] = useState(false);

  if (pendingUser) {
    const authFirmsList = firms.filter(f => pendingUser.authorizedFirms.includes(f.id));
    return (
      <div className="login-page-bg">
        <div className="login-box" style={{ width: 380 }}>
          <div className="window-header">
            <div className="window-title">Logo J-Plus | Çalışma Firması Seçimi</div>
          </div>
          <div className="window-content" style={{ padding: '20px', background: '#f8f9fa', textAlign: 'center' }}>
            <div style={{ marginBottom: 15, fontSize: 12, color: '#333', fontWeight: 600 }}>Giriş yapmak istediğiniz Kurum / Firmayı seçiniz:</div>
            {authFirmsList.length === 0 ? (
              <div style={{ color: '#c92a2a', fontSize: 12, marginBottom: 15 }}>Yetkili olduğunuz hiçbir firma bulunamadı.<br/>Lütfen Sistem İşletmenine (admin) başvurunuz.</div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 20 }}>
                {authFirmsList.map(f => (
                  <button key={f.id} className="btn-classic" style={{ padding: '10px', justifyContent: 'space-between', fontWeight: 'bold' }} onClick={() => onFirmSelect(f.id)}>
                    <span>({f.id}) {f.name}</span>
                    <span className="status-badge badge-info">{f.period || '2026'}</span>
                  </button>
                ))}
              </div>
            )}
            <button className="btn-classic" onClick={() => onFirmSelect(null)}>Geri Dön</button>
          </div>
        </div>
      </div>
    );
  }

  const handleLogin = (e) => {
    e.preventDefault();
    if (username && password) onLogin(username, password);
    else alert("Lütfen kullanıcı adı ve şifre giriniz.");
  };

  return (
    <div className="login-page-bg">
      <div className="login-box">
        <div className="login-header-brand">
          <h1>OZN J-Plus ERP</h1>
          <p>Kurumsal Muhasebe ve Kaynak Planlama Sistemi</p>
        </div>
        <div className="login-content">
          <form onSubmit={handleLogin}>
            <div style={{ fontSize: 11.5, color: '#2c3e50', marginBottom: 15, background: '#e2e6ea', padding: 10, borderRadius: 3, borderLeft: '4px solid #0f2c59' }}>
              <div style={{marginBottom: 4}}><strong>Hızlı Giriş Hesapları:</strong></div>
              <div style={{display: 'flex', justifyContent: 'space-between'}}>
                <span>Müşteri/Muhasebe: <b>veli</b> / <b>123</b></span>
                <span style={{cursor: 'pointer', color: '#0056b3', textDecoration: 'underline'}} onClick={() => {setUsername('veli'); setPassword('123');}}>Doldur</span>
              </div>
              <div style={{display: 'flex', justifyContent: 'space-between', marginTop: 2}}>
                <span>Sistem İşletmeni (SYS): <b>admin</b> / <b>123</b></span>
                <span style={{cursor: 'pointer', color: '#0056b3', textDecoration: 'underline'}} onClick={() => {setUsername('admin'); setPassword('123');}}>Doldur</span>
              </div>
            </div>
            <div className="login-input-group">
              <label className="login-label">Kullanıcı Kodu</label>
              <div className="login-input-wrapper">
                <div className="login-icon"><User size={15} /></div>
                <input type="text" className="login-input-field" value={username} onChange={e => setUsername(e.target.value)} autoFocus />
              </div>
            </div>
            <div className="login-input-group">
              <label className="login-label">Parola</label>
              <div className="login-input-wrapper">
                <div className="login-icon"><Lock size={15} /></div>
                <input type={showPassword ? "text" : "password"} className="login-input-field" value={password} onChange={e => setPassword(e.target.value)} />
                <div className="login-icon" style={{ cursor: 'pointer' }} onClick={() => setShowPassword(!showPassword)}>
                  {showPassword ? <EyeOff size={15} /> : <Eye size={15} />}
                </div>
              </div>
            </div>
            <button type="submit" className="login-btn-submit">Sisteme Bağlan (F2)</button>
          </form>
        </div>
        <div className="login-footer">OZN J-Plus Kurumsal ERP v4.5 &copy; 2026</div>
      </div>
    </div>
  );
};

// ==========================================
// 1. SİSTEM İŞLETMENİ (SYS - Admin)
// ==========================================
const SaaSAdminApp = ({ session, onLogout, users, setUsers, firms, setFirms }) => {
  const [tabs, setTabs] = useState([{ id: 'firmalar', title: 'Firmalar ve Dönemler' }]);
  const [activeTab, setActiveTab] = useState('firmalar');
  const [openFirmModal, setOpenFirmModal] = useState(false);
  const [openUserModal, setOpenUserModal] = useState(false);
  const [openAuthModal, setOpenAuthModal] = useState(false);
  const [selectedFirm, setSelectedFirm] = useState(null);
  const [selectedUser, setSelectedUser] = useState(null);
  
  const [newFirmCode, setNewFirmCode] = useState('');
  const [newFirmName, setNewFirmName] = useState('');
  const [newFirmTaxOffice, setNewFirmTaxOffice] = useState('');
  const [newFirmTaxNo, setNewFirmTaxNo] = useState('');
  
  const [newUsername, setNewUsername] = useState('');
  const [newPassword, setNewPassword] = useState('123');
  const [newUserFirmId, setNewUserFirmId] = useState('');

  const handleCreateFirm = () => {
    if(!newFirmCode || !newFirmName) return alert("Firma No ve Unvan giriniz.");
    if(firms.find(f => f.id === newFirmCode)) return alert("Bu firma numarası kullanımda!");
    setFirms([...firms, { 
      id: newFirmCode, 
      name: newFirmName.toUpperCase(), 
      period: '01 (2026)',
      taxOffice: newFirmTaxOffice || 'Merkez V.D.',
      taxNo: newFirmTaxNo || '1111111111',
      createdAt: new Date().toLocaleDateString('tr-TR') 
    }]);
    setOpenFirmModal(false); setNewFirmCode(''); setNewFirmName(''); setNewFirmTaxOffice(''); setNewFirmTaxNo('');
  };

  const handleCreateUser = () => {
    if(!newUsername || !newPassword) return alert("Kullanıcı bilgilerini giriniz.");
    if(users.find(u => u.username === newUsername)) return alert("Kullanıcı sistemde mevcut!");
    if(!newUserFirmId) return alert("Lütfen kullanıcıyı atayacağınız firmayı seçin!");
    setUsers([...users, { username: newUsername, password: newPassword, role: 'user', authorizedFirms: [newUserFirmId] }]);
    setOpenUserModal(false); setNewUsername(''); setNewPassword('123'); setNewUserFirmId('');
  };

  const handleToggleAuth = (firmId) => {
    const userToUpdate = users.find(u => u.username === selectedUser);
    if (!userToUpdate) return;
    let updatedAuths = [...userToUpdate.authorizedFirms];
    if (updatedAuths.includes(firmId)) updatedAuths = updatedAuths.filter(id => id !== firmId);
    else updatedAuths.push(firmId);
    setUsers(users.map(u => u.username === selectedUser ? { ...u, authorizedFirms: updatedAuths } : u));
  };

  const closeTab = (e, id) => {
    e.stopPropagation();
    const newTabs = tabs.filter(t => t.id !== id);
    setTabs(newTabs);
    if (activeTab === id) setActiveTab(newTabs.length > 0 ? newTabs[newTabs.length - 1].id : null);
  };

  return (
    <div className="app-container">
      {openFirmModal && (
        <div className="modal-overlay">
          <div className="classic-window" style={{ width: '460px' }}>
            <div className="window-header"><div className="window-title">Yeni Firma (Kurum) Tanımla</div><X size={14} className="window-close" onClick={() => setOpenFirmModal(false)} /></div>
            <div className="window-content">
              <div className="form-group"><label className="form-label">Firma No:</label><input type="text" className="form-input" value={newFirmCode} onChange={e=>setNewFirmCode(e.target.value)} autoFocus placeholder="Örn: 002" maxLength={3}/></div>
              <div className="form-group"><label className="form-label">Firma Unvanı:</label><input type="text" className="form-input" value={newFirmName} onChange={e=>setNewFirmName(e.target.value)} placeholder="Örn: ANADOLU LOJİSTİK A.Ş."/></div>
              <div className="form-group"><label className="form-label">Vergi Dairesi:</label><input type="text" className="form-input" value={newFirmTaxOffice} onChange={e=>setNewFirmTaxOffice(e.target.value)} placeholder="Örn: Marmara Kurumlar"/></div>
              <div className="form-group"><label className="form-label">Vergi Numarası:</label><input type="text" className="form-input" value={newFirmTaxNo} onChange={e=>setNewFirmTaxNo(e.target.value)} maxLength={11} placeholder="10 Haneli VKN"/></div>
            </div>
            <div className="window-footer"><button className="btn-classic primary" onClick={handleCreateFirm}>Kaydet (F2)</button><button className="btn-classic" onClick={() => setOpenFirmModal(false)}>Vazgeç</button></div>
          </div>
        </div>
      )}

      {openUserModal && (
        <div className="modal-overlay">
          <div className="classic-window" style={{ width: '440px' }}>
            <div className="window-header"><div className="window-title">Yeni Kullanıcı Kartı ve Firma Ataması</div><X size={14} className="window-close" onClick={() => setOpenUserModal(false)} /></div>
            <div className="window-content">
              <div className="form-group"><label className="form-label">Kullanıcı Kodu:</label><input type="text" className="form-input" value={newUsername} onChange={e=>setNewUsername(e.target.value)} autoFocus /></div>
              <div className="form-group"><label className="form-label">Giriş Şifresi:</label><input type="text" className="form-input" value={newPassword} onChange={e=>setNewPassword(e.target.value)} /></div>
              <div style={{height: 1, backgroundColor: '#ccc', margin: '10px 0'}}></div>
              <div className="form-group"><label className="form-label">Yetkili Firma:</label>
                <select className="form-select" value={newUserFirmId} onChange={e => setNewUserFirmId(e.target.value)}>
                  <option value="">-- Atanacağı Firmayı Seçin --</option>
                  {firms.map(f => <option key={f.id} value={f.id}>({f.id}) {f.name}</option>)}
                </select>
              </div>
            </div>
            <div className="window-footer"><button className="btn-classic primary" onClick={handleCreateUser}>Kaydet (F2)</button><button className="btn-classic" onClick={() => setOpenUserModal(false)}>Vazgeç</button></div>
          </div>
        </div>
      )}

      {openAuthModal && selectedUser && (
        <div className="modal-overlay">
          <div className="classic-window" style={{ width: '480px', height: '380px' }}>
            <div className="window-header"><div className="window-title">Firma Çalışma Yetkileri - [{selectedUser}]</div><X size={14} className="window-close" onClick={() => setOpenAuthModal(false)} /></div>
            <div className="window-content" style={{ padding: 0 }}>
              <table className="data-grid">
                <thead><tr><th style={{width: 50, textAlign: 'center'}}>Yetki</th><th style={{width: 80}}>Firma No</th><th>Firma Unvanı</th></tr></thead>
                <tbody>
                  {firms.map(f => {
                    const usr = users.find(u => u.username === selectedUser);
                    const isAuth = usr?.authorizedFirms.includes(f.id);
                    return (
                      <tr key={f.id} onClick={() => handleToggleAuth(f.id)} style={{cursor: 'pointer'}}>
                        <td style={{textAlign: 'center'}}><input type="checkbox" checked={isAuth} readOnly style={{cursor: 'pointer'}} /></td>
                        <td><b>{f.id}</b></td><td>{f.name}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div className="window-footer"><button className="btn-classic primary" onClick={() => setOpenAuthModal(false)}>Tamam</button></div>
          </div>
        </div>
      )}

      <header className="app-header">
        <div className="brand-title"><span className="brand-badge">SYS</span> OZN J-Plus Sistem İşletmeni (Yönetim Konsolu)</div>
        <div className="header-tools">
          <span style={{ color: '#ffd700', fontWeight: 600 }}>YÖNETİCİ: {session.username.toUpperCase()}</span>
          <span onClick={onLogout} style={{color: '#ff8787', fontWeight: 700}}>Oturumu Kapat</span>
        </div>
      </header>

      <div className="workspace">
        <aside className="sidebar">
          <div className="sidebar-header">Sistem Ağacı</div>
          <div className="tree-view">
            {adminMenuData.map(folder => (
              <div key={folder.id}>
                <div className="tree-node"><div className="tree-icon"><FolderOpen size={15} fill="#f39c12" color="#d35400"/></div><span className="tree-text" style={{fontWeight: 600}}>{folder.title}</span></div>
                <div className="tree-children">
                  {folder.children.map(item => (
                    <div key={item.id} className={`tree-node ${activeTab === item.id ? 'active-node' : ''}`} onClick={() => { if(!tabs.find(t=>t.id===item.id)) setTabs([...tabs, {id:item.id, title:item.title}]); setActiveTab(item.id); }}>
                      <div className="tree-icon"><FileText size={13} color="#0056b3" /></div><span className="tree-text">{item.title}</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </aside>
        <main className="content-area">
          <div className="tabs-container">
            {tabs.map(tab => (
              <div key={tab.id} className={`tab ${activeTab === tab.id ? 'active' : ''}`} onClick={() => setActiveTab(tab.id)}>
                {tab.title} <X size={13} className="tab-close" onClick={(e) => closeTab(e, tab.id)} />
              </div>
            ))}
          </div>
          <div className="module-container">
            {activeTab === 'firmalar' && (
              <>
                <div className="toolbar">
                  <button className="toolbar-btn" onClick={() => setOpenFirmModal(true)}><Plus size={15} color="#2b8a3e" /> Yeni Firma Aç</button>
                </div>
                <div className="grid-container">
                  <table className="data-grid">
                    <thead><tr><th style={{width: 30}}></th><th style={{width: 80}}>Firma No</th><th>Firma Unvanı</th><th style={{width: 140}}>Vergi Dairesi</th><th style={{width: 120}}>Vergi No</th><th style={{width: 110}}>Aktif Dönem</th><th style={{width: 110}}>Kayıt Tarihi</th></tr></thead>
                    <tbody>
                      {firms.map(f => (
                        <tr key={f.id} className={selectedFirm === f.id ? 'selected' : ''} onClick={() => setSelectedFirm(f.id)}>
                          <td style={{textAlign: 'center'}}>{selectedFirm === f.id ? '►' : ''}</td>
                          <td><b>{f.id}</b></td><td>{f.name}</td><td>{f.taxOffice || 'Merkez V.D.'}</td><td>{f.taxNo || '6480912345'}</td><td>{f.period || '01 (2026)'}</td><td>{f.createdAt}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
            {activeTab === 'kullanicilar' && (
              <>
                <div className="toolbar">
                  <button className="toolbar-btn" onClick={() => setOpenUserModal(true)}><Plus size={15} color="#2b8a3e" /> Yeni Kullanıcı Ekle</button>
                  <div className="toolbar-sep"></div>
                  <button className="toolbar-btn" onClick={() => { if(!selectedUser) return alert("Lütfen listeden bir kullanıcı seçiniz!"); setOpenAuthModal(true); }}>
                    <Key size={15} color="#0056b3" /> Firma Yetkileri (Atama Yap)
                  </button>
                </div>
                <div className="grid-container">
                  <table className="data-grid">
                    <thead><tr><th style={{width: 30}}></th><th style={{width: 150}}>Kullanıcı Kodu</th><th style={{width: 120}}>Rol / Yetki</th><th>Atanmış Firmalar</th></tr></thead>
                    <tbody>
                      {users.map(u => (
                        <tr key={u.username} className={selectedUser === u.username ? 'selected' : ''} onClick={() => setSelectedUser(u.username)}>
                          <td style={{textAlign: 'center'}}>{selectedUser === u.username ? '►' : ''}</td>
                          <td><b>{u.username}</b></td>
                          <td><span className={`status-badge ${u.role === 'admin' ? 'badge-borc' : 'badge-info'}`}>{u.role.toUpperCase()}</span></td>
                          <td>{u.role === 'admin' ? 'TÜM FİRMALAR (SİSTEM İŞLETMENİ)' : u.authorizedFirms.map(fid => `(${fid}) ${firms.find(x=>x.id===fid)?.name || ''}`).join(', ') || 'Firma Atanmamış'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        </main>
      </div>
      <footer className="status-bar">
        <div>Sistem İşletmeni Hazır | Toplam Firma: {firms.length} | Toplam Kullanıcı: {users.length}</div>
        <div>OZN J-Plus SYS v4.5</div>
      </footer>
    </div>
  );
};


// ==========================================
// 2. ERP CLIENT APP (Logo J-Plus Tam Kapsamlı Muhasebe)
// ==========================================
const ERPClientApp = ({ session, onLogout, firmDetails }) => {
  const [treeData, setTreeData] = useState(erpMenuData);
  const [tabs, setTabs] = useState([
    { id: 'cari_hesaplar', title: 'Cari Hesaplar' },
    { id: 'malzemeler', title: 'Malzemeler (Stok Kartları)' },
    { id: 'satis_faturalari', title: 'Satış Faturaları' }
  ]);
  const [activeTab, setActiveTab] = useState('cari_hesaplar');
  const [openModal, setOpenModal] = useState(null);
  const [editItem, setEditItem] = useState(null);
  const [selectedGridItem, setSelectedGridItem] = useState(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [ekstreTarget, setEkstreTarget] = useState(null);

  // --- CORE ERP STATE ---
  const [cariler, setCariler] = useState([]);
  const [stoklar, setStoklar] = useState([]);
  const [malzemeFisleri, setMalzemeFisleri] = useState([]);
  const [siparisler, setSiparisler] = useState([]);
  const [faturalar, setFaturalar] = useState([]);
  const [kasalar, setKasalar] = useState([]);
  const [bankalar, setBankalar] = useState([]);
  const [finansFisleri, setFinansFisleri] = useState([]); // Kasa & Banka Fişleri
  const [cariFisleri, setCariFisleri] = useState([]);
  const [cekSenetler, setCekSenetler] = useState([]);
  const [hesapPlani, setHesapPlani] = useState(defaultHesapPlani);
  const [muhasebeFisleri, setMuhasebeFisleri] = useState([]); // Yevmiye / Mahsup Fişleri

  const fId = session.firmaId;

  // Load state from local storage
  useEffect(() => {
    const load = (key, setter, fallback) => {
      const v = localStorage.getItem(`ozn_${key}_${fId}`);
      if (v) {
        try { setter(JSON.parse(v)); } catch(e) { if (fallback) setter(fallback); }
      } else if (fallback) {
        setter(fallback);
      }
    };
    load('cariler', setCariler, []);
    load('stoklar', setStoklar, []);
    load('malzemeFisleri', setMalzemeFisleri, []);
    load('siparisler', setSiparisler, []);
    load('faturalar', setFaturalar, []);
    load('kasalar', setKasalar, []);
    load('bankalar', setBankalar, []);
    load('finansFisleri', setFinansFisleri, []);
    load('cariFisleri', setCariFisleri, []);
    load('cekSenetler', setCekSenetler, []);
    load('hesapPlani', setHesapPlani, defaultHesapPlani);
    load('muhasebeFisleri', setMuhasebeFisleri, []);
  }, [fId]);

  // Save state to local storage
  useEffect(() => {
    const save = (key, data) => localStorage.setItem(`ozn_${key}_${fId}`, JSON.stringify(data));
    save('cariler', cariler);
    save('stoklar', stoklar);
    save('malzemeFisleri', malzemeFisleri);
    save('siparisler', siparisler);
    save('faturalar', faturalar);
    save('kasalar', kasalar);
    save('bankalar', bankalar);
    save('finansFisleri', finansFisleri);
    save('cariFisleri', cariFisleri);
    save('cekSenetler', cekSenetler);
    save('hesapPlani', hesapPlani);
    save('muhasebeFisleri', muhasebeFisleri);
  }, [cariler, stoklar, malzemeFisleri, siparisler, faturalar, kasalar, bankalar, finansFisleri, cariFisleri, cekSenetler, hesapPlani, muhasebeFisleri, fId]);

  // Helper: Otomatik Muhasebe (Yevmiye/Mahsup) Entegrasyonu
  const postToGeneralLedger = (fisTipi, aciklama, lines, currentHesapPlani = hesapPlani, currentMuhFisleri = muhasebeFisleri) => {
    // lines: [{ accountCode: '120', borc: 1200, alacak: 0, desc: '...' }, ...]
    const totalBorc = lines.reduce((s, l) => s + (parseFloat(l.borc) || 0), 0);
    const totalAlacak = lines.reduce((s, l) => s + (parseFloat(l.alacak) || 0), 0);
    const newFis = {
      id: Date.now() + Math.floor(Math.random() * 1000),
      code: `MHS-${String(currentMuhFisleri.length + 1).padStart(4, '0')}`,
      date: new Date().toISOString().split('T')[0],
      type: fisTipi,
      desc: aciklama,
      totalBorc,
      totalAlacak,
      lines
    };
    const updatedPlani = currentHesapPlani.map(acc => {
      const matching = lines.filter(l => l.accountCode === acc.code);
      if (matching.length === 0) return acc;
      const addBorc = matching.reduce((s, l) => s + (parseFloat(l.borc) || 0), 0);
      const addAlacak = matching.reduce((s, l) => s + (parseFloat(l.alacak) || 0), 0);
      return {
        ...acc,
        borc: (parseFloat(acc.borc) || 0) + addBorc,
        alacak: (parseFloat(acc.alacak) || 0) + addAlacak
      };
    });
    setHesapPlani(updatedPlani);
    setMuhasebeFisleri([...currentMuhFisleri, newFis]);
  };

  // --- ÖRNEK KURUMSAL VERİ YÜKLEME (DEMO DATA) ---
  const handleLoadDemoData = () => {
    if (cariler.length > 0 && !window.confirm("Mevcut verilere örnek kurumsal muhasebe kayıtları eklensin mi? (Sıfırdan örnek şirket verisi yüklenecektir)")) return;
    
    const demoCariler = [
      { id: 101, code: '120.01.001', desc: 'YILDIZ TEKNOLOJİ VE BİLİŞİM A.Ş.', type: 'Alıcı', taxOffice: 'Maslak V.D.', taxNo: '9876543210', phone: '0212 345 67 89', city: 'İstanbul', riskLimit: 500000, balance: 74000 },
      { id: 102, code: '120.01.002', desc: 'MEGA İNŞAAT YAPI MARKET LTD. ŞTİ.', type: 'Alıcı', taxOffice: 'Kadıköy V.D.', taxNo: '6543210987', phone: '0216 444 11 22', city: 'İstanbul', riskLimit: 250000, balance: 25800 },
      { id: 103, code: '320.01.001', desc: 'GLOBAL DONANIM İTHALAT A.Ş.', type: 'Satıcı', taxOffice: 'Şişli V.D.', taxNo: '1122334455', phone: '0212 222 33 44', city: 'İstanbul', riskLimit: 1000000, balance: -82000 },
      { id: 104, code: '320.01.002', desc: 'ANADOLU OFİS MOBİLYALARI SAN. TİC.', type: 'Satıcı', taxOffice: 'Ostim V.D.', taxNo: '5566778899', phone: '0312 385 00 11', city: 'Ankara', riskLimit: 300000, balance: -18000 }
    ];

    const demoStoklar = [
      { id: 201, code: 'STK.001', desc: 'LENOVO THINKPAD KURUMSAL NOTEBOOK', type: 'Ticari Mal', unit: 'Adet', vat: 20, buyPrice: 30000, sellPrice: 40000, minStock: 5, balance: 12 },
      { id: 202, code: 'STK.002', desc: 'DELL 27" IPS PROFESYONEL MONİTÖR', type: 'Ticari Mal', unit: 'Adet', vat: 20, buyPrice: 6500, sellPrice: 9000, minStock: 10, balance: 25 },
      { id: 203, code: 'STK.003', desc: 'LOGITECH KABLOSUZ KLAVYE & MOUSE SET', type: 'Ticari Mal', unit: 'Adet', vat: 20, buyPrice: 1200, sellPrice: 1850, minStock: 15, balance: 40 },
      { id: 204, code: 'STK.004', desc: 'CAT6 UTP NETWORK KABLOSU (305M KUTU)', type: 'Sarf Malzemesi', unit: 'Kutu', vat: 20, buyPrice: 2800, sellPrice: 4200, minStock: 5, balance: 8 }
    ];

    const demoKasalar = [
      { id: 301, code: 'KASA.01', desc: 'MERKEZ TL NAKİT KASASI', currency: 'TRY', balance: 145000 },
      { id: 302, code: 'KASA.02', desc: 'ŞUBE GÜNLÜK GİDER KASASI', currency: 'TRY', balance: 22500 }
    ];

    const demoBankalar = [
      { id: 401, code: 'BNK.01', desc: 'GARANTİ BBVA - MERKEZ TİCARİ HESAP', branch: 'Levent Ticari', iban: 'TR45 0006 2000 1230 0006 2988 11', balance: 680000 },
      { id: 402, code: 'BNK.02', desc: 'TÜRKİYE İŞ BANKASI - KURUMSAL', branch: 'Maslak', iban: 'TR12 0006 4000 0011 2345 6789 00', balance: 315000 }
    ];

    const demoFaturalar = [
      { id: 501, code: 'ALF-2026001', date: '2026-08-10', cariId: 103, type: 'Satınalma Faturası', subTotal: 150000, discountTotal: 0, vatTotal: 30000, total: 180000, lines: [{ id: 1, stokId: 201, qty: 5, price: 30000, discount: 0, vat: 20, total: 180000 }] },
      { id: 502, code: 'STF-2026001', date: '2026-08-12', cariId: 101, type: 'Satış Faturası', subTotal: 100000, discountTotal: 0, vatTotal: 20000, total: 120000, lines: [{ id: 1, stokId: 201, qty: 2, price: 40000, discount: 0, vat: 20, total: 96000 }, { id: 2, stokId: 202, qty: 2, price: 10000, discount: 0, vat: 20, total: 24000 }] }
    ];

    const demoFinans = [
      { id: 601, code: 'BNK-0001', date: '2026-08-14', arac: 'Banka', aracId: 401, cariId: 101, type: 'Gelen Havale / Tahsilat', total: 46000, desc: 'Yıldız Teknoloji Fatura Mahsubu' },
      { id: 602, code: 'KAS-0001', date: '2026-08-15', arac: 'Kasa', aracId: 301, cariId: 103, type: 'Tediye (Ödeme) Fişi', total: 98000, desc: 'Global Donanım Nakit Ödeme' }
    ];

    const demoSiparisler = [
      { id: 701, code: 'SIP-00101', date: '2026-08-16', cariId: 102, type: 'Satış Siparişi', status: 'Onaylandı (Bekliyor)', total: 54000, lines: [{ id: 1, stokId: 202, qty: 5, price: 9000, discount: 0, vat: 20, total: 54000 }] }
    ];

    const demoCekler = [
      { id: 801, code: 'CEK-9001', vade: '2026-09-30', cariId: 101, type: 'Müşteri Çeki (Alınan)', bankName: 'Yapı Kredi Bankası', checkNo: '7845120', status: 'Portföyde', total: 50000 }
    ];

    const demoPlani = [
      { code: '100', desc: 'KASA HESABI', type: 'Aktif', borc: 265500, alacak: 98000 },
      { code: '101', desc: 'ALINAN ÇEKLER', type: 'Aktif', borc: 50000, alacak: 0 },
      { code: '102', desc: 'BANKALAR', type: 'Aktif', borc: 995000, alacak: 0 },
      { code: '103', desc: 'VERİLEN ÇEKLER VE ÖDEME EMİRLERİ (-)', type: 'Aktif', borc: 0, alacak: 0 },
      { code: '120', desc: 'ALICILAR', type: 'Aktif', borc: 195800, alacak: 96000 },
      { code: '121', desc: 'ALACAK SENETLERİ', type: 'Aktif', borc: 0, alacak: 0 },
      { code: '153', desc: 'TİCARİ MALLAR', type: 'Aktif', borc: 580000, alacak: 0 },
      { code: '191', desc: 'İNDİRİLECEK KDV', type: 'Aktif', borc: 30000, alacak: 0 },
      { code: '320', desc: 'SATICILAR', type: 'Pasif', borc: 98000, alacak: 198000 },
      { code: '321', desc: 'BORÇ SENETLERİ', type: 'Pasif', borc: 0, alacak: 0 },
      { code: '391', desc: 'HESAPLANAN KDV', type: 'Pasif', borc: 0, alacak: 20000 },
      { code: '500', desc: 'SERMAYE', type: 'Pasif', borc: 0, alacak: 1702300 },
      { code: '600', desc: 'YURTİÇİ SATIŞLAR', type: 'Gelir', borc: 0, alacak: 100000 },
      { code: '610', desc: 'SATIŞTAN İADELER (-)', type: 'Gider', borc: 0, alacak: 0 },
      { code: '621', desc: 'SATILAN TİCARİ MALLAR MALİYETİ (-)', type: 'Gider', borc: 0, alacak: 0 },
      { code: '770', desc: 'GENEL YÖNETİM GİDERLERİ', type: 'Gider', borc: 0, alacak: 0 }
    ];

    const demoMuhFisleri = [
      { id: 901, code: 'MHS-0001', date: '2026-08-10', type: 'Mahsup Fişi', desc: 'ALF-2026001 Nolu Satınalma Faturası Muhasebe Kaydı', totalBorc: 180000, totalAlacak: 180000, lines: [
        { accountCode: '153', borc: 150000, alacak: 0, desc: 'Ticari Mal Alımı' },
        { accountCode: '191', borc: 30000, alacak: 0, desc: '%20 İndirilecek KDV' },
        { accountCode: '320', borc: 0, alacak: 180000, desc: 'Global Donanım A.Ş.' }
      ]},
      { id: 902, code: 'MHS-0002', date: '2026-08-12', type: 'Mahsup Fişi', desc: 'STF-2026001 Nolu Satış Faturası Muhasebe Kaydı', totalBorc: 120000, totalAlacak: 120000, lines: [
        { accountCode: '120', borc: 120000, alacak: 0, desc: 'Yıldız Teknoloji A.Ş.' },
        { accountCode: '600', borc: 0, alacak: 100000, desc: 'Yurtiçi Satışlar' },
        { accountCode: '391', borc: 0, alacak: 20000, desc: '%20 Hesaplanan KDV' }
      ]}
    ];

    setCariler(demoCariler);
    setStoklar(demoStoklar);
    setKasalar(demoKasalar);
    setBankalar(demoBankalar);
    setFaturalar(demoFaturalar);
    setFinansFisleri(demoFinans);
    setSiparisler(demoSiparisler);
    setCekSenetler(demoCekler);
    setHesapPlani(demoPlani);
    setMuhasebeFisleri(demoMuhFisleri);
  };

  const toggleFolder = (id) => {
    const newTree = JSON.parse(JSON.stringify(treeData));
    const findAndToggle = (nodes) => {
      for (let node of nodes) {
        if (node.id === id) { node.isOpen = !node.isOpen; return true; }
        if (node.children && findAndToggle(node.children)) return true;
      }
      return false;
    };
    findAndToggle(newTree); setTreeData(newTree);
  };

  const handleMenuSelect = (node) => {
    if (node.type === 'folder') return;
    if (!tabs.find(t => t.id === node.id)) setTabs([...tabs, { id: node.id, title: node.title }]);
    setActiveTab(node.id);
    setSelectedGridItem(null);
    setSearchTerm('');
  };

  const closeTab = (e, id) => {
    e.stopPropagation();
    const newTabs = tabs.filter(t => t.id !== id);
    setTabs(newTabs);
    if (activeTab === id) setActiveTab(newTabs.length > 0 ? newTabs[newTabs.length - 1].id : null);
  };

  // --- SAVE HANDLERS WITH FULL ACCOUNTING INTEGRATION ---
  const handleSaveCari = (data) => {
    if (editItem) {
      setCariler(cariler.map(c => c.id === editItem.id ? { ...c, ...data } : c));
    } else {
      setCariler([...cariler, { ...data, id: Date.now(), balance: parseFloat(data.balance) || 0 }]);
    }
    setEditItem(null);
    setOpenModal(null);
  };

  const handleSaveStok = (data) => {
    if (editItem) {
      setStoklar(stoklar.map(s => s.id === editItem.id ? { ...s, ...data } : s));
    } else {
      setStoklar([...stoklar, { ...data, id: Date.now(), balance: parseFloat(data.balance) || 0 }]);
    }
    setEditItem(null);
    setOpenModal(null);
  };

  const handleSaveMalzemeFisi = (data) => {
    setMalzemeFisleri([...malzemeFisleri, { ...data, id: Date.now() }]);
    const qty = parseFloat(data.qty) || 0;
    const isGiris = data.type.includes('Giriş') || data.type.includes('Fazlası') || data.type.includes('Devir');
    const degisim = isGiris ? qty : -qty;
    setStoklar(stoklar.map(s => s.id.toString() === data.stokId.toString() ? { ...s, balance: (parseFloat(s.balance) || 0) + degisim } : s));
    setOpenModal(null);
  };

  const handleSaveKasaBanka = (data, type) => {
    if (type === 'Kasa') {
      if (editItem) setKasalar(kasalar.map(k => k.id === editItem.id ? { ...k, ...data } : k));
      else setKasalar([...kasalar, { ...data, id: Date.now(), balance: parseFloat(data.balance) || 0 }]);
    }
    if (type === 'Banka') {
      if (editItem) setBankalar(bankalar.map(b => b.id === editItem.id ? { ...b, ...data } : b));
      else setBankalar([...bankalar, { ...data, id: Date.now(), balance: parseFloat(data.balance) || 0 }]);
    }
    setEditItem(null);
    setOpenModal(null);
  };

  const handleSaveSiparis = (data) => {
    setSiparisler([...siparisler, { ...data, id: Date.now(), status: 'Onaylandı (Bekliyor)' }]);
    setOpenModal(null);
  };

  const handleConvertOrderToInvoice = (order) => {
    if (order.status === 'Faturalandı') return alert("Bu sipariş zaten faturalandırılmış!");
    const faturaTipi = order.type === 'Satınalma Siparişi' ? 'Satınalma Faturası' : 'Satış Faturası';
    const subTotal = order.lines.reduce((s, l) => s + (l.qty * l.price), 0);
    const vatTotal = order.lines.reduce((s, l) => s + (l.qty * l.price * ((l.vat || 20)/100)), 0);
    const total = parseFloat(order.total) || (subTotal + vatTotal);

    handleSaveFatura({
      code: `${faturaTipi === 'Satınalma Faturası' ? 'ALF' : 'STF'}-${Math.floor(10000 + Math.random()*90000)}`,
      date: new Date().toISOString().split('T')[0],
      cariId: order.cariId,
      type: faturaTipi,
      subTotal,
      discountTotal: 0,
      vatTotal,
      total,
      lines: order.lines
    });

    setSiparisler(siparisler.map(s => s.id === order.id ? { ...s, status: 'Faturalandı' } : s));
    alert(`${order.code} nolu sipariş başarıyla "${faturaTipi}" olarak faturalandırıldı ve stok/cari/muhasebe hesaplarına işlendi!`);
  };

  const handleSaveFatura = (data) => {
    setFaturalar([...faturalar, { ...data, id: Date.now() }]);
    
    const tutar = parseFloat(data.total) || 0;
    const subTotal = parseFloat(data.subTotal) || tutar;
    const vatTotal = parseFloat(data.vatTotal) || 0;
    const isAlim = data.type === 'Satınalma Faturası' || data.type === 'Satış İade Faturası';
    
    // 1. Cari Bakiye Güncelle
    const bakiyeDegisimiCari = isAlim ? -tutar : tutar;
    const targetCari = cariler.find(c => c.id.toString() === data.cariId.toString());
    setCariler(cariler.map(c => c.id.toString() === data.cariId.toString() ? { ...c, balance: (parseFloat(c.balance) || 0) + bakiyeDegisimiCari } : c));

    // 2. Stok Bakiyelerini Güncelle
    let updatedStoklar = [...stoklar];
    data.lines.forEach(line => {
      if (line.stokId) {
        const qty = parseFloat(line.qty) || 0;
        const qtyDegisimi = isAlim ? qty : -qty;
        updatedStoklar = updatedStoklar.map(s => s.id.toString() === line.stokId.toString() ? { ...s, balance: (parseFloat(s.balance) || 0) + qtyDegisimi } : s);
      }
    });
    setStoklar(updatedStoklar);

    // 3. Otomatik Muhasebeleştirme (Yevmiye Fişi)
    if (isAlim) {
      postToGeneralLedger('Mahsup Fişi', `${data.code} - ${targetCari?.desc || ''} Satınalma Faturası`, [
        { accountCode: '153', borc: subTotal, alacak: 0, desc: 'Ticari Mal Alımı' },
        { accountCode: '191', borc: vatTotal, alacak: 0, desc: 'İndirilecek KDV' },
        { accountCode: '320', borc: 0, alacak: tutar, desc: targetCari?.desc || 'Satıcılar' }
      ]);
    } else {
      postToGeneralLedger('Mahsup Fişi', `${data.code} - ${targetCari?.desc || ''} Satış Faturası`, [
        { accountCode: '120', borc: tutar, alacak: 0, desc: targetCari?.desc || 'Alıcılar' },
        { accountCode: '600', borc: 0, alacak: subTotal, desc: 'Yurtiçi Satışlar' },
        { accountCode: '391', borc: 0, alacak: vatTotal, desc: 'Hesaplanan KDV' }
      ]);
    }

    setOpenModal(null);
  };

  const handleSaveFinansFisi = (data) => {
    setFinansFisleri([...finansFisleri, { ...data, id: Date.now(), date: data.date || new Date().toISOString().split('T')[0] }]);
    const tutar = parseFloat(data.total) || 0;
    const isTahsilat = data.type.includes('Tahsilat') || data.type.includes('Gelen');
    
    // 1. Cari Güncelle
    const cariDegisim = isTahsilat ? -tutar : tutar;
    const targetCari = cariler.find(c => c.id.toString() === data.cariId.toString());
    setCariler(cariler.map(c => c.id.toString() === data.cariId.toString() ? { ...c, balance: (parseFloat(c.balance) || 0) + cariDegisim } : c));

    // 2. Kasa/Banka Güncelle
    const aracDegisim = isTahsilat ? tutar : -tutar;
    if (data.arac === 'Kasa') {
      setKasalar(kasalar.map(k => k.id.toString() === data.aracId.toString() ? { ...k, balance: (parseFloat(k.balance) || 0) + aracDegisim } : k));
    } else if (data.arac === 'Banka') {
      setBankalar(bankalar.map(b => b.id.toString() === data.aracId.toString() ? { ...b, balance: (parseFloat(b.balance) || 0) + aracDegisim } : b));
    }

    // 3. Muhasebe Entegrasyonu
    const kasaBankaKod = data.arac === 'Kasa' ? '100' : '102';
    const cariMuhKod = isTahsilat ? '120' : '320';
    if (isTahsilat) {
      postToGeneralLedger('Tahsil Fişi', `${data.code} - ${targetCari?.desc || ''} Tahsilat`, [
        { accountCode: kasaBankaKod, borc: tutar, alacak: 0, desc: `${data.arac} Tahsilatı` },
        { accountCode: cariMuhKod, borc: 0, alacak: tutar, desc: targetCari?.desc || 'Cari Tahsilat' }
      ]);
    } else {
      postToGeneralLedger('Tediye Fişi', `${data.code} - ${targetCari?.desc || ''} Ödeme`, [
        { accountCode: cariMuhKod, borc: tutar, alacak: 0, desc: targetCari?.desc || 'Cari Ödeme' },
        { accountCode: kasaBankaKod, borc: 0, alacak: tutar, desc: `${data.arac} Ödemesi` }
      ]);
    }

    setOpenModal(null);
  };

  const handleSaveCariFisi = (data) => {
    setCariFisleri([...cariFisleri, { ...data, id: Date.now() }]);
    const tutar = parseFloat(data.total) || 0;
    const cariDegisim = data.type === 'Alacak Dekontu' ? -tutar : tutar;
    setCariler(cariler.map(c => c.id.toString() === data.cariId.toString() ? { ...c, balance: (parseFloat(c.balance) || 0) + cariDegisim } : c));
    setOpenModal(null);
  };

  const handleSaveCekSenet = (data) => {
    setCekSenetler([...cekSenetler, { ...data, id: Date.now(), status: data.type.includes('Müşteri') ? 'Portföyde' : 'Tedavülde (Verildi)' }]);
    const tutar = parseFloat(data.total) || 0;
    const isAlinan = data.type.includes('Müşteri');
    const cariDegisim = isAlinan ? -tutar : tutar;
    const targetCari = cariler.find(c => c.id.toString() === data.cariId.toString());
    setCariler(cariler.map(c => c.id.toString() === data.cariId.toString() ? { ...c, balance: (parseFloat(c.balance) || 0) + cariDegisim } : c));

    if (isAlinan) {
      postToGeneralLedger('Mahsup Fişi', `${data.code} - Müşteri Çek/Senet Giriş Bordrosu`, [
        { accountCode: '101', borc: tutar, alacak: 0, desc: `${data.bankName} - ${data.checkNo}` },
        { accountCode: '120', borc: 0, alacak: tutar, desc: targetCari?.desc || 'Alıcılar' }
      ]);
    } else {
      postToGeneralLedger('Mahsup Fişi', `${data.code} - Kendi Çek/Senet Çıkış Bordrosu`, [
        { accountCode: '320', borc: tutar, alacak: 0, desc: targetCari?.desc || 'Satıcılar' },
        { accountCode: '103', borc: 0, alacak: tutar, desc: `${data.bankName} - ${data.checkNo}` }
      ]);
    }
    setOpenModal(null);
  };

  const handleSaveMuhasebeFisi = (data) => {
    postToGeneralLedger(data.type, data.desc, data.lines);
    setOpenModal(null);
  };

  // --- DELETE & EXCEL EXPORT ---
  const handleDeleteSelected = () => {
    if (!selectedGridItem) return alert("Lütfen silmek istediğiniz satırı seçiniz!");
    if (!window.confirm("Seçili kaydı silmek istediğinize emin misiniz?")) return;
    if (activeTab === 'cari_hesaplar') setCariler(cariler.filter(x => x.id !== selectedGridItem));
    else if (activeTab === 'malzemeler') setStoklar(stoklar.filter(x => x.id !== selectedGridItem));
    else if (activeTab === 'kasa_kartlari') setKasalar(kasalar.filter(x => x.id !== selectedGridItem));
    else if (activeTab === 'banka_kartlari') setBankalar(bankalar.filter(x => x.id !== selectedGridItem));
    else if (activeTab.includes('siparisleri')) setSiparisler(siparisler.filter(x => x.id !== selectedGridItem));
    else if (activeTab.includes('faturalari')) setFaturalar(faturalar.filter(x => x.id !== selectedGridItem));
    else if (activeTab === 'cek_senet') setCekSenetler(cekSenetler.filter(x => x.id !== selectedGridItem));
    else if (activeTab === 'cari_hesap_fisleri') setCariFisleri(cariFisleri.filter(x => x.id !== selectedGridItem));
    else if (activeTab === 'muhasebe_fisleri') setMuhasebeFisleri(muhasebeFisleri.filter(x => x.id !== selectedGridItem));
    setSelectedGridItem(null);
  };

  const handleEditSelected = () => {
    if (!selectedGridItem) return alert("Lütfen değiştirmek istediğiniz kartı seçiniz!");
    if (activeTab === 'cari_hesaplar') {
      const item = cariler.find(x => x.id === selectedGridItem);
      if (item) { setEditItem(item); setOpenModal('cari'); }
    } else if (activeTab === 'malzemeler') {
      const item = stoklar.find(x => x.id === selectedGridItem);
      if (item) { setEditItem(item); setOpenModal('stok'); }
    } else if (activeTab === 'kasa_kartlari') {
      const item = kasalar.find(x => x.id === selectedGridItem);
      if (item) { setEditItem(item); setOpenModal('kasa'); }
    } else if (activeTab === 'banka_kartlari') {
      const item = bankalar.find(x => x.id === selectedGridItem);
      if (item) { setEditItem(item); setOpenModal('banka'); }
    } else {
      alert("Fiş ve fatura kayıtları muhasebe standartları gereği 'İncele' ile görüntülenebilir.");
    }
  };

  const handleOpenEkstre = () => {
    if (activeTab === 'cari_hesaplar') {
      const item = cariler.find(x => x.id === selectedGridItem) || cariler[0];
      if (!item) return alert("Cari hesap bulunamadı!");
      setEkstreTarget({ type: 'cari', item });
    } else if (activeTab === 'malzemeler') {
      const item = stoklar.find(x => x.id === selectedGridItem) || stoklar[0];
      if (!item) return alert("Malzeme kartı bulunamadı!");
      setEkstreTarget({ type: 'stok', item });
    } else {
      alert("Ekstre / Hareket dökümü için Cari Hesaplar veya Malzemeler sekmesinde bir kayıt seçiniz.");
    }
  };

  const handleExportExcel = () => {
    let exportData = [];
    if (activeTab === 'cari_hesaplar') exportData = cariler;
    else if (activeTab === 'malzemeler') exportData = stoklar;
    else if (activeTab.includes('faturalari')) exportData = faturalar;
    else if (activeTab === 'hesap_plani' || activeTab === 'mizan_raporu') exportData = hesapPlani;
    else if (activeTab === 'cek_senet') exportData = cekSenetler;
    else exportData = muhasebeFisleri;

    if (!exportData || exportData.length === 0) return alert("Dışa aktarılacak veri bulunamadı!");
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeTab.substring(0, 25));
    XLSX.writeFile(wb, `OZN_JPlus_${activeTab}_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const TreeNode = ({ node }) => {
    const isFolder = node.type === 'folder';
    return (
      <div>
        <div className={`tree-node ${activeTab === node.id ? 'active-node' : ''}`} onClick={() => isFolder ? toggleFolder(node.id) : handleMenuSelect(node)}>
          <div className="tree-icon">
            {isFolder ? (node.isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />) : <span style={{ width: 13, display: 'inline-block' }}></span>}
            {isFolder ? (node.isOpen ? <FolderOpen size={15} fill="#f39c12" color="#d35400"/> : <Folder size={15} fill="#f39c12" color="#d35400"/>) : <FileText size={13} color="#0056b3" />}
          </div>
          <span className="tree-text">{node.title}</span>
        </div>
        {isFolder && node.isOpen && node.children && <div className="tree-children">{node.children.map(child => <TreeNode key={child.id} node={child} />)}</div>}
      </div>
    );
  };

  const getToolbarConfig = () => {
    switch(activeTab) {
      case 'cari_hesaplar': return { label: 'Yeni Cari Kart (F2)', action: () => { setEditItem(null); setOpenModal('cari'); } };
      case 'cari_hesap_fisleri': return { label: 'Cari Hesap Fişi Kes', action: () => setOpenModal('cari_fisi') };
      case 'malzemeler': return { label: 'Yeni Malzeme Kartı', action: () => { setEditItem(null); setOpenModal('stok'); } };
      case 'malzeme_fisleri': return { label: 'Malzeme Fişi (Sayım/Sarf)', action: () => setOpenModal('malzeme_fisi') };
      case 'satinalma_siparisleri': return { label: 'Satınalma Siparişi Gir', action: () => setOpenModal('siparis_alim') };
      case 'satis_siparisleri': return { label: 'Satış Siparişi Gir', action: () => setOpenModal('siparis_satis') };
      case 'satinalma_faturalari': return { label: 'Satınalma Faturası Kes', action: () => setOpenModal('fatura_alim') };
      case 'satis_faturalari': return { label: 'Satış Faturası Kes', action: () => setOpenModal('fatura_satis') };
      case 'kasa_kartlari': return { label: 'Yeni Kasa Kartı', action: () => { setEditItem(null); setOpenModal('kasa'); } };
      case 'kasa_islemleri': return { label: 'Kasa Tahsilat / Ödeme', action: () => setOpenModal('kasa_islem') };
      case 'banka_kartlari': return { label: 'Yeni Banka Hesabı', action: () => { setEditItem(null); setOpenModal('banka'); } };
      case 'banka_islemleri': return { label: 'Banka Havale / EFT Fişi', action: () => setOpenModal('banka_islem') };
      case 'cek_senet': return { label: 'Çek / Senet Bordrosu Gir', action: () => setOpenModal('cek_senet') };
      case 'hesap_plani': return { label: 'Yeni Muhasebe Hesabı', action: () => setOpenModal('hesap_ekle') };
      case 'muhasebe_fisleri': return { label: 'Yeni Mahsup / Yevmiye Fişi', action: () => setOpenModal('muhasebe_fisi') };
      default: return { label: 'Yeni Kayıt', action: null };
    }
  };

  // --- KPI SUMMARY METRICS ---
  const totalAlacak = useMemo(() => cariler.filter(c => c.balance > 0).reduce((s, c) => s + parseFloat(c.balance), 0), [cariler]);
  const totalBorc = useMemo(() => cariler.filter(c => c.balance < 0).reduce((s, c) => s + Math.abs(parseFloat(c.balance)), 0), [cariler]);
  const totalKasaBanka = useMemo(() => kasalar.reduce((s, k) => s + parseFloat(k.balance || 0), 0) + bankalar.reduce((s, b) => s + parseFloat(b.balance || 0), 0), [kasalar, bankalar]);
  const totalStokDegeri = useMemo(() => stoklar.reduce((s, st) => s + ((parseFloat(st.balance) || 0) * (parseFloat(st.buyPrice) || 0)), 0), [stoklar]);

  const filterRow = (text) => !searchTerm || (text && text.toString().toLowerCase().includes(searchTerm.toLowerCase()));

  return (
    <div className="app-container">
      {/* MODALS */}
      {openModal === 'cari' && <CariHesapModal onClose={() => { setOpenModal(null); setEditItem(null); }} onSave={handleSaveCari} initialData={editItem} />}
      {openModal === 'cari_fisi' && <CariFisiModal onClose={() => setOpenModal(null)} cariler={cariler} onSave={handleSaveCariFisi} />}
      {openModal === 'stok' && <StokModal onClose={() => { setOpenModal(null); setEditItem(null); }} onSave={handleSaveStok} initialData={editItem} />}
      {openModal === 'malzeme_fisi' && <MalzemeFisiModal onClose={() => setOpenModal(null)} stoklar={stoklar} onSave={handleSaveMalzemeFisi} />}
      {openModal === 'kasa' && <BankaKasaModal onClose={() => { setOpenModal(null); setEditItem(null); }} onSave={(d) => handleSaveKasaBanka(d, 'Kasa')} type="Kasa" initialData={editItem} />}
      {openModal === 'banka' && <BankaKasaModal onClose={() => { setOpenModal(null); setEditItem(null); }} onSave={(d) => handleSaveKasaBanka(d, 'Banka')} type="Banka" initialData={editItem} />}
      
      {(openModal === 'siparis_alim' || openModal === 'siparis_satis') && (
        <FaturaModal 
          onClose={() => setOpenModal(null)} 
          cariler={cariler} stoklar={stoklar}
          faturaTipi={openModal === 'siparis_alim' ? 'Satınalma Siparişi' : 'Satış Siparişi'}
          isOrder={true}
          onSave={handleSaveSiparis} 
        />
      )}

      {(openModal === 'fatura_alim' || openModal === 'fatura_satis') && (
        <FaturaModal 
          onClose={() => setOpenModal(null)} 
          cariler={cariler} stoklar={stoklar}
          faturaTipi={openModal === 'fatura_alim' ? 'Satınalma Faturası' : 'Satış Faturası'}
          onSave={handleSaveFatura} 
        />
      )}

      {(openModal === 'kasa_islem' || openModal === 'banka_islem') && (
        <FinansFisiModal 
          onClose={() => setOpenModal(null)}
          cariler={cariler} aracListesi={openModal === 'kasa_islem' ? kasalar : bankalar}
          arac={openModal === 'kasa_islem' ? 'Kasa' : 'Banka'}
          onSave={handleSaveFinansFisi}
        />
      )}

      {openModal === 'cek_senet' && (
        <CekSenetModal onClose={() => setOpenModal(null)} cariler={cariler} onSave={handleSaveCekSenet} />
      )}

      {openModal === 'hesap_ekle' && (
        <HesapEkleModal onClose={() => setOpenModal(null)} onSave={(acc) => { setHesapPlani([...hesapPlani, { ...acc, borc: 0, alacak: 0 }].sort((a,b)=>a.code.localeCompare(b.code))); setOpenModal(null); }} />
      )}

      {openModal === 'muhasebe_fisi' && (
        <MuhasebeFisiModal onClose={() => setOpenModal(null)} hesapPlani={hesapPlani} onSave={handleSaveMuhasebeFisi} />
      )}

      {ekstreTarget && (
        <EkstreModal 
          target={ekstreTarget} 
          onClose={() => setEkstreTarget(null)} 
          faturalar={faturalar} 
          finansFisleri={finansFisleri} 
          cariFisleri={cariFisleri}
          cekSenetler={cekSenetler}
          malzemeFisleri={malzemeFisleri}
        />
      )}

      {/* TOP HEADER */}
      <header className="app-header">
        <div className="brand-title">
          <span className="brand-badge">J-PLUS</span>
          OZN Kurumsal ERP
          <span style={{fontSize: 12, fontWeight: 'normal', color: '#cbd5e1', marginLeft: 6}}>
            | Kurum: ({firmDetails?.id}) {firmDetails?.name} | Dönem: {firmDetails?.period || '01 (2026)'}
          </span>
        </div>
        <div className="header-tools">
          <button className="btn-classic" style={{padding: '2px 8px', fontSize: 11, background: '#ffd43b', borderColor: '#f08c00'}} onClick={handleLoadDemoData}>
            <RefreshCw size={12} /> Örnek Şirket Verisi Yükle
          </button>
          <span style={{ color: '#fff' }}><User size={13} style={{verticalAlign: 'text-bottom', marginRight: 4}}/>{session.username.toUpperCase()}</span>
          <span onClick={onLogout} style={{color: '#ffc107', fontWeight: 700}}>Çıkış</span>
        </div>
      </header>

      {/* KPI SUMMARY STRIP */}
      <div className="kpi-grid">
        <div className="kpi-card" style={{borderLeftColor: '#2b8a3e'}}>
          <span className="kpi-title">Toplam Ticari Alacaklar (120)</span>
          <span className="kpi-val" style={{color: '#2b8a3e'}}>{formatMoney(totalAlacak)} TL</span>
        </div>
        <div className="kpi-card" style={{borderLeftColor: '#c92a2a'}}>
          <span className="kpi-title">Toplam Ticari Borçlar (320)</span>
          <span className="kpi-val" style={{color: '#c92a2a'}}>{formatMoney(totalBorc)} TL</span>
        </div>
        <div className="kpi-card" style={{borderLeftColor: '#0056b3'}}>
          <span className="kpi-title">Hazır Değerler (Kasa + Banka)</span>
          <span className="kpi-val" style={{color: '#0056b3'}}>{formatMoney(totalKasaBanka)} TL</span>
        </div>
        <div className="kpi-card" style={{borderLeftColor: '#e67700'}}>
          <span className="kpi-title">Envanter (Stok) Maliyet Değeri</span>
          <span className="kpi-val" style={{color: '#d9480f'}}>{formatMoney(totalStokDegeri)} TL</span>
        </div>
      </div>

      <div className="workspace">
        {/* LEFT TREE MENU */}
        <aside className="sidebar">
          <div className="sidebar-header">
            <span>Modül Ağacı (J-Plus)</span>
          </div>
          <div className="tree-view">{treeData.map(node => <TreeNode key={node.id} node={node} />)}</div>
        </aside>

        {/* MDI CONTENT AREA */}
        <main className="content-area">
          <div className="tabs-container">
            {tabs.map(tab => (
              <div key={tab.id} className={`tab ${activeTab === tab.id ? 'active' : ''}`} onClick={() => { setActiveTab(tab.id); setSelectedGridItem(null); }}>
                {tab.title} <X size={13} className="tab-close" onClick={(e) => closeTab(e, tab.id)} />
              </div>
            ))}
          </div>

          {activeTab ? (
            <div className="module-container">
              <div className="toolbar">
                {getToolbarConfig().action && (
                  <button className="toolbar-btn" onClick={getToolbarConfig().action}>
                    <Plus size={15} color="#2b8a3e" /> {getToolbarConfig().label}
                  </button>
                )}
                <button className="toolbar-btn" onClick={handleEditSelected}><Edit size={14} color="#0056b3" /> Değiştir</button>
                <button className="toolbar-btn" onClick={handleDeleteSelected}><Trash2 size={14} color="#c92a2a" /> Çıkar (Sil)</button>
                <div className="toolbar-sep"></div>
                {(activeTab === 'cari_hesaplar' || activeTab === 'malzemeler') && (
                  <button className="toolbar-btn" onClick={handleOpenEkstre}>
                    <BookOpen size={14} color="#6741d9" /> Ekstre / Hareketler (F5)
                  </button>
                )}
                <button className="toolbar-btn" onClick={handleExportExcel}>
                  <FileSpreadsheet size={14} color="#2b8a3e" /> Excel'e Aktar
                </button>
                <button className="toolbar-btn" onClick={() => window.print()}>
                  <Printer size={14} color="#495057" /> Yazdır
                </button>

                <div className="toolbar-search">
                  <Search size={14} color="#57606f" />
                  <input 
                    type="text" 
                    placeholder="Kod veya açıklama ile filtrele..." 
                    value={searchTerm} 
                    onChange={e => setSearchTerm(e.target.value)} 
                  />
                </div>
              </div>

              <div className="grid-container">
                {/* 1. CARİ HESAPLAR */}
                {activeTab === 'cari_hesaplar' && (
                  <table className="data-grid">
                    <thead>
                      <tr>
                        <th style={{width: 28}}></th>
                        <th style={{width: 120}}>Cari Kodu</th>
                        <th>Cari Hesap Unvanı</th>
                        <th style={{width: 90, textAlign: 'center'}}>Türü</th>
                        <th style={{width: 130}}>Vergi Dairesi</th>
                        <th style={{width: 110}}>VKN / TCKN</th>
                        <th style={{width: 110}}>Şehir / Tel</th>
                        <th style={{width: 130, textAlign: 'right'}}>Bakiye (TL)</th>
                        <th style={{width: 75, textAlign: 'center'}}>B/A</th>
                      </tr>
                    </thead>
                    <tbody>
                      {cariler.filter(r => filterRow(r.code) || filterRow(r.desc)).map(row => {
                        const bal = parseFloat(row.balance) || 0;
                        return (
                          <tr key={row.id} className={selectedGridItem === row.id ? 'selected' : ''} onClick={() => setSelectedGridItem(row.id)} onDoubleClick={() => setEkstreTarget({ type: 'cari', item: row })}>
                            <td style={{textAlign: 'center'}}>{selectedGridItem === row.id ? '►' : ''}</td>
                            <td><b>{row.code}</b></td>
                            <td>{row.desc}</td>
                            <td style={{textAlign: 'center'}}><span className="status-badge badge-info">{row.type}</span></td>
                            <td>{row.taxOffice || '-'}</td>
                            <td>{row.taxNo || '-'}</td>
                            <td>{row.city || '-'}</td>
                            <td style={{textAlign: 'right', fontWeight: 700, color: bal < 0 ? '#c92a2a' : bal > 0 ? '#2b8a3e' : '#333'}}>
                              {formatMoney(Math.abs(bal))}
                            </td>
                            <td style={{textAlign: 'center'}}>
                              <span className={`status-badge ${bal > 0 ? 'badge-alacak' : bal < 0 ? 'badge-borc' : ''}`}>
                                {bal > 0 ? 'BORÇLU (B)' : bal < 0 ? 'ALACAKLI (A)' : '-'}
                              </span>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}

                {/* 2. MALZEMELER (STOK KARTLARI) */}
                {activeTab === 'malzemeler' && (
                  <table className="data-grid">
                    <thead>
                      <tr>
                        <th style={{width: 28}}></th>
                        <th style={{width: 110}}>Malzeme Kodu</th>
                        <th>Malzeme Açıklaması</th>
                        <th style={{width: 110, textAlign: 'center'}}>Kart Türü</th>
                        <th style={{width: 70, textAlign: 'center'}}>Birim</th>
                        <th style={{width: 70, textAlign: 'center'}}>KDV %</th>
                        <th style={{width: 115, textAlign: 'right'}}>Alış Fiyatı</th>
                        <th style={{width: 115, textAlign: 'right'}}>Satış Fiyatı</th>
                        <th style={{width: 110, textAlign: 'right'}}>Fiili Stok</th>
                      </tr>
                    </thead>
                    <tbody>
                      {stoklar.filter(r => filterRow(r.code) || filterRow(r.desc)).map(row => (
                        <tr key={row.id} className={selectedGridItem === row.id ? 'selected' : ''} onClick={() => setSelectedGridItem(row.id)} onDoubleClick={() => setEkstreTarget({ type: 'stok', item: row })}>
                          <td style={{textAlign: 'center'}}>{selectedGridItem === row.id ? '►' : ''}</td>
                          <td><b>{row.code}</b></td>
                          <td>{row.desc}</td>
                          <td style={{textAlign: 'center'}}>{row.type}</td>
                          <td style={{textAlign: 'center'}}>{row.unit}</td>
                          <td style={{textAlign: 'center'}}>%{row.vat || 20}</td>
                          <td style={{textAlign: 'right'}}>{formatMoney(row.buyPrice || 0)} TL</td>
                          <td style={{textAlign: 'right'}}>{formatMoney(row.sellPrice || 0)} TL</td>
                          <td style={{textAlign: 'right', fontWeight: 700, color: (row.balance || 0) <= (row.minStock || 0) ? '#c92a2a' : '#0f2c59'}}>
                            {row.balance} {row.unit}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {/* 3. MALZEME YÖNETİM FİŞLERİ */}
                {activeTab === 'malzeme_fisleri' && (
                  <table className="data-grid">
                    <thead>
                      <tr>
                        <th style={{width: 28}}></th>
                        <th style={{width: 120}}>Fiş No</th>
                        <th style={{width: 100}}>Tarih</th>
                        <th style={{width: 160}}>Fiş Türü</th>
                        <th>Malzeme Açıklaması</th>
                        <th style={{width: 110, textAlign: 'right'}}>Miktar</th>
                        <th>Açıklama</th>
                      </tr>
                    </thead>
                    <tbody>
                      {malzemeFisleri.map(row => {
                        const st = stoklar.find(s => s.id.toString() === row.stokId.toString());
                        return (
                          <tr key={row.id} className={selectedGridItem === row.id ? 'selected' : ''} onClick={() => setSelectedGridItem(row.id)}>
                            <td style={{textAlign: 'center'}}>{selectedGridItem === row.id ? '►' : ''}</td>
                            <td><b>{row.code}</b></td>
                            <td>{row.date}</td>
                            <td><span className="status-badge badge-info">{row.type}</span></td>
                            <td>{st ? `${st.code} - ${st.desc}` : '-'}</td>
                            <td style={{textAlign: 'right', fontWeight: 700}}>{row.qty} {st?.unit || 'Adet'}</td>
                            <td>{row.desc || '-'}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}

                {/* 4. SİPARİŞLER (SATINALMA & SATIŞ) */}
                {(activeTab === 'satinalma_siparisleri' || activeTab === 'satis_siparisleri') && (
                  <table className="data-grid">
                    <thead>
                      <tr>
                        <th style={{width: 28}}></th>
                        <th style={{width: 120}}>Sipariş No</th>
                        <th style={{width: 100}}>Tarih</th>
                        <th>Cari Hesap Unvanı</th>
                        <th style={{width: 140, textAlign: 'center'}}>Durumu</th>
                        <th style={{width: 130, textAlign: 'right'}}>Sipariş Tutarı</th>
                        <th style={{width: 150, textAlign: 'center'}}>İşlem</th>
                      </tr>
                    </thead>
                    <tbody>
                      {siparisler.filter(s => activeTab === 'satinalma_siparisleri' ? s.type === 'Satınalma Siparişi' : s.type === 'Satış Siparişi').map(row => (
                        <tr key={row.id} className={selectedGridItem === row.id ? 'selected' : ''} onClick={() => setSelectedGridItem(row.id)}>
                          <td style={{textAlign: 'center'}}>{selectedGridItem === row.id ? '►' : ''}</td>
                          <td><b>{row.code}</b></td>
                          <td>{row.date}</td>
                          <td>{cariler.find(c => c.id.toString() === row.cariId.toString())?.desc || '-'}</td>
                          <td style={{textAlign: 'center'}}>
                            <span className={`status-badge ${row.status === 'Faturalandı' ? 'badge-alacak' : 'badge-warn'}`}>{row.status}</span>
                          </td>
                          <td style={{textAlign: 'right', fontWeight: 700}}>{formatMoney(row.total)} TL</td>
                          <td style={{textAlign: 'center'}}>
                            {row.status !== 'Faturalandı' && (
                              <button className="btn-classic primary" style={{padding: '2px 8px', fontSize: 11}} onClick={(e) => { e.stopPropagation(); handleConvertOrderToInvoice(row); }}>
                                <ArrowRightLeft size={11} /> Faturaya Aktar
                              </button>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {/* 5. FATURALAR (SATINALMA & SATIŞ) */}
                {(activeTab === 'satinalma_faturalari' || activeTab === 'satis_faturalari') && (
                  <table className="data-grid">
                    <thead>
                      <tr>
                        <th style={{width: 28}}></th>
                        <th style={{width: 125}}>Fatura No</th>
                        <th style={{width: 100}}>Tarih</th>
                        <th>Cari Hesap Unvanı</th>
                        <th style={{width: 140, textAlign: 'center'}}>Fatura Türü</th>
                        <th style={{width: 120, textAlign: 'right'}}>Matrah (Ara Top.)</th>
                        <th style={{width: 110, textAlign: 'right'}}>KDV Tutarı</th>
                        <th style={{width: 130, textAlign: 'right'}}>Net Genel Toplam</th>
                      </tr>
                    </thead>
                    <tbody>
                      {faturalar.filter(f => activeTab === 'satinalma_faturalari' ? f.type === 'Satınalma Faturası' : f.type === 'Satış Faturası').map(row => (
                        <tr key={row.id} className={selectedGridItem === row.id ? 'selected' : ''} onClick={() => setSelectedGridItem(row.id)}>
                          <td style={{textAlign: 'center'}}>{selectedGridItem === row.id ? '►' : ''}</td>
                          <td><b>{row.code}</b></td>
                          <td>{row.date}</td>
                          <td>{cariler.find(c => c.id.toString() === row.cariId.toString())?.desc || 'Bilinmiyor'}</td>
                          <td style={{textAlign: 'center'}}><span className="status-badge badge-info">{row.type}</span></td>
                          <td style={{textAlign: 'right'}}>{formatMoney(row.subTotal || row.total)} TL</td>
                          <td style={{textAlign: 'right'}}>{formatMoney(row.vatTotal || 0)} TL</td>
                          <td style={{textAlign: 'right', fontWeight: 700, color: '#0f2c59'}}>{formatMoney(row.total)} TL</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {/* 6. CARİ HESAP FİŞLERİ */}
                {activeTab === 'cari_hesap_fisleri' && (
                  <table className="data-grid">
                    <thead>
                      <tr>
                        <th style={{width: 28}}></th>
                        <th style={{width: 120}}>Fiş No</th>
                        <th style={{width: 100}}>Tarih</th>
                        <th style={{width: 150}}>Fiş Türü</th>
                        <th>Cari Hesap Unvanı</th>
                        <th>Açıklama</th>
                        <th style={{width: 130, textAlign: 'right'}}>Tutar (TL)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {cariFisleri.map(row => (
                        <tr key={row.id} className={selectedGridItem === row.id ? 'selected' : ''} onClick={() => setSelectedGridItem(row.id)}>
                          <td style={{textAlign: 'center'}}>{selectedGridItem === row.id ? '►' : ''}</td>
                          <td><b>{row.code}</b></td>
                          <td>{row.date}</td>
                          <td><span className="status-badge badge-info">{row.type}</span></td>
                          <td>{cariler.find(c => c.id.toString() === row.cariId.toString())?.desc || '-'}</td>
                          <td>{row.desc || '-'}</td>
                          <td style={{textAlign: 'right', fontWeight: 700}}>{formatMoney(row.total)} TL</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {/* 7. KASA VE BANKA KARTLARI */}
                {(activeTab === 'kasa_kartlari' || activeTab === 'banka_kartlari') && (
                  <table className="data-grid">
                    <thead>
                      <tr>
                        <th style={{width: 28}}></th>
                        <th style={{width: 120}}>Kodu</th>
                        <th>Hesap / Kasa Açıklaması</th>
                        <th>Şube / IBAN Detayı</th>
                        <th style={{width: 90, textAlign: 'center'}}>Döviz</th>
                        <th style={{width: 140, textAlign: 'right'}}>Bakiye</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(activeTab === 'kasa_kartlari' ? kasalar : bankalar).map(row => (
                        <tr key={row.id} className={selectedGridItem === row.id ? 'selected' : ''} onClick={() => setSelectedGridItem(row.id)}>
                          <td style={{textAlign: 'center'}}>{selectedGridItem === row.id ? '►' : ''}</td>
                          <td><b>{row.code}</b></td>
                          <td>{row.desc}</td>
                          <td>{row.iban ? `${row.branch || ''} - ${row.iban}` : 'Merkez Kasa'}</td>
                          <td style={{textAlign: 'center'}}>{row.currency || 'TRY'}</td>
                          <td style={{textAlign: 'right', fontWeight: 700, color: '#2b8a3e'}}>{formatMoney(row.balance)} TL</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {/* 8. KASA VE BANKA İŞLEMLERİ */}
                {(activeTab === 'kasa_islemleri' || activeTab === 'banka_islemleri') && (
                  <table className="data-grid">
                    <thead>
                      <tr>
                        <th style={{width: 28}}></th>
                        <th style={{width: 110}}>Fiş No</th>
                        <th style={{width: 100}}>Tarih</th>
                        <th style={{width: 180}}>İşlem Türü</th>
                        <th>Kasa / Banka</th>
                        <th>Karşı Cari Hesap</th>
                        <th>Açıklama</th>
                        <th style={{width: 130, textAlign: 'right'}}>Tutar (TL)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {finansFisleri.filter(f => activeTab === 'kasa_islemleri' ? f.arac === 'Kasa' : f.arac === 'Banka').map(row => {
                        const aracList = row.arac === 'Kasa' ? kasalar : bankalar;
                        const aracObj = aracList.find(a => a.id.toString() === row.aracId?.toString());
                        return (
                          <tr key={row.id} className={selectedGridItem === row.id ? 'selected' : ''} onClick={() => setSelectedGridItem(row.id)}>
                            <td style={{textAlign: 'center'}}>{selectedGridItem === row.id ? '►' : ''}</td>
                            <td><b>{row.code}</b></td>
                            <td>{row.date}</td>
                            <td><span className={`status-badge ${row.type.includes('Tahsilat') || row.type.includes('Gelen') ? 'badge-alacak' : 'badge-borc'}`}>{row.type}</span></td>
                            <td>{aracObj?.desc || row.arac}</td>
                            <td>{cariler.find(c => c.id.toString() === row.cariId.toString())?.desc || '-'}</td>
                            <td>{row.desc || '-'}</td>
                            <td style={{textAlign: 'right', fontWeight: 700}}>{formatMoney(row.total)} TL</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}

                {/* 9. ÇEK VE SENET BORDROLARI */}
                {activeTab === 'cek_senet' && (
                  <table className="data-grid">
                    <thead>
                      <tr>
                        <th style={{width: 28}}></th>
                        <th style={{width: 110}}>Bordro No</th>
                        <th style={{width: 100}}>Vade Tarihi</th>
                        <th style={{width: 170}}>Bordro / Evrak Türü</th>
                        <th>Cari Hesap</th>
                        <th>Banka / Çek No</th>
                        <th style={{width: 120, textAlign: 'center'}}>Durumu</th>
                        <th style={{width: 130, textAlign: 'right'}}>Tutar (TL)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {cekSenetler.map(row => (
                        <tr key={row.id} className={selectedGridItem === row.id ? 'selected' : ''} onClick={() => setSelectedGridItem(row.id)}>
                          <td style={{textAlign: 'center'}}>{selectedGridItem === row.id ? '►' : ''}</td>
                          <td><b>{row.code}</b></td>
                          <td><b>{row.vade}</b></td>
                          <td><span className="status-badge badge-info">{row.type}</span></td>
                          <td>{cariler.find(c => c.id.toString() === row.cariId.toString())?.desc || '-'}</td>
                          <td>{row.bankName} (#{row.checkNo})</td>
                          <td style={{textAlign: 'center'}}><span className="status-badge badge-alacak">{row.status}</span></td>
                          <td style={{textAlign: 'right', fontWeight: 700}}>{formatMoney(row.total)} TL</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {/* 10. GENEL MUHASEBE HESAP PLANI & MİZAN */}
                {(activeTab === 'hesap_plani' || activeTab === 'mizan_raporu') && (
                  <table className="data-grid">
                    <thead>
                      <tr>
                        <th style={{width: 28}}></th>
                        <th style={{width: 110}}>Hesap Kodu</th>
                        <th>Muhasebe Hesap Adı (Tek Düzen Hesap Planı)</th>
                        <th style={{width: 100, textAlign: 'center'}}>Hesap Tipi</th>
                        <th style={{width: 140, textAlign: 'right'}}>Borç Toplamı (TL)</th>
                        <th style={{width: 140, textAlign: 'right'}}>Alacak Toplamı (TL)</th>
                        <th style={{width: 140, textAlign: 'right'}}>Borç Bakiyesi</th>
                        <th style={{width: 140, textAlign: 'right'}}>Alacak Bakiyesi</th>
                      </tr>
                    </thead>
                    <tbody>
                      {hesapPlani.filter(r => filterRow(r.code) || filterRow(r.desc)).map(row => {
                        const b = parseFloat(row.borc) || 0;
                        const a = parseFloat(row.alacak) || 0;
                        const bakiyeBorc = b > a ? b - a : 0;
                        const bakiyeAlacak = a > b ? a - b : 0;
                        return (
                          <tr key={row.code} className={selectedGridItem === row.code ? 'selected' : ''} onClick={() => setSelectedGridItem(row.code)}>
                            <td style={{textAlign: 'center'}}>{selectedGridItem === row.code ? '►' : ''}</td>
                            <td><b>{row.code}</b></td>
                            <td style={{fontWeight: 600}}>{row.desc}</td>
                            <td style={{textAlign: 'center'}}><span className="status-badge badge-info">{row.type || 'Ana Hesap'}</span></td>
                            <td style={{textAlign: 'right'}}>{formatMoney(b)}</td>
                            <td style={{textAlign: 'right'}}>{formatMoney(a)}</td>
                            <td style={{textAlign: 'right', fontWeight: 700, color: '#0056b3'}}>{bakiyeBorc > 0 ? formatMoney(bakiyeBorc) : '-'}</td>
                            <td style={{textAlign: 'right', fontWeight: 700, color: '#2b8a3e'}}>{bakiyeAlacak > 0 ? formatMoney(bakiyeAlacak) : '-'}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}

                {/* 11. MUHASEBE FİŞLERİ (MAHSUP / YEVMİYE DEFTERİ) */}
                {activeTab === 'muhasebe_fisleri' && (
                  <table className="data-grid">
                    <thead>
                      <tr>
                        <th style={{width: 28}}></th>
                        <th style={{width: 115}}>Fiş No</th>
                        <th style={{width: 100}}>Tarih</th>
                        <th style={{width: 120}}>Fiş Türü</th>
                        <th>Yevmiye / Fiş Açıklaması ve Detay Satırları</th>
                        <th style={{width: 130, textAlign: 'right'}}>Toplam Borç</th>
                        <th style={{width: 130, textAlign: 'right'}}>Toplam Alacak</th>
                      </tr>
                    </thead>
                    <tbody>
                      {muhasebeFisleri.map(row => (
                        <tr key={row.id} className={selectedGridItem === row.id ? 'selected' : ''} onClick={() => setSelectedGridItem(row.id)}>
                          <td style={{textAlign: 'center'}}>{selectedGridItem === row.id ? '►' : ''}</td>
                          <td><b>{row.code}</b></td>
                          <td>{row.date}</td>
                          <td><span className="status-badge badge-info">{row.type}</span></td>
                          <td>
                            <div style={{fontWeight: 600}}>{row.desc}</div>
                            <div style={{fontSize: 11, color: '#57606f'}}>
                              {row.lines?.map((l, i) => `${l.accountCode} (${l.borc > 0 ? 'B:'+formatMoney(l.borc) : 'A:'+formatMoney(l.alacak)})`).join('  |  ')}
                            </div>
                          </td>
                          <td style={{textAlign: 'right', fontWeight: 700, color: '#0056b3'}}>{formatMoney(row.totalBorc)} TL</td>
                          <td style={{textAlign: 'right', fontWeight: 700, color: '#2b8a3e'}}>{formatMoney(row.totalAlacak)} TL</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>

              {/* BOTTOM MODULE SUMMARY */}
              <div className="grid-summary-bar">
                <div>Listelenen Kayıt Sayısı: <b>{
                  activeTab === 'cari_hesaplar' ? cariler.length :
                  activeTab === 'malzemeler' ? stoklar.length :
                  activeTab.includes('faturalari') ? faturalar.length :
                  activeTab === 'muhasebe_fisleri' ? muhasebeFisleri.length :
                  hesapPlani.length
                }</b></div>
                <div style={{color: '#0f2c59'}}>
                  Çift tıklayarak veya <b>Ekstre / Hareketler</b> butonuyla hesap detaylarını inceleyebilirsiniz.
                </div>
              </div>
            </div>
          ) : (
            <div style={{flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#57606f'}}>
              Lütfen sol menü ağacından çalışmak istediğiniz modülü seçiniz.
            </div>
          )}
        </main>
      </div>

      <footer className="status-bar">
        <div>Durum: <b>Çevrimiçi / Entegre Muhasebe Aktif</b> | Aktif Pencere: {activeTab}</div>
        <div style={{color: '#0f2c59', fontWeight: 'bold'}}>Kurum: ({firmDetails?.id}) {firmDetails?.name} | VKN: {firmDetails?.taxNo || '6480912345'}</div>
      </footer>
    </div>
  );
};


// ==========================================
// 3. MAIN APP CONTROLLER
// ==========================================
function App() {
  const [session, setSession] = useState(null); 
  const [users, setUsers] = useState(initialUsers);
  const [firms, setFirms] = useState(initialFirms);
  const [pendingUser, setPendingUser] = useState(null);

  useEffect(() => {
    const sUsers = localStorage.getItem('sys_users');
    const sFirms = localStorage.getItem('sys_firms');
    if (sUsers) { try { setUsers(JSON.parse(sUsers)); } catch(e) {} }
    if (sFirms) { try { setFirms(JSON.parse(sFirms)); } catch(e) {} }

    const sSession = sessionStorage.getItem('ozn_session');
    if (sSession) { try { setSession(JSON.parse(sSession)); } catch(e) {} }
  }, []);

  useEffect(() => {
    localStorage.setItem('sys_users', JSON.stringify(users));
    localStorage.setItem('sys_firms', JSON.stringify(firms));
  }, [users, firms]);

  const handleLogin = (username, password) => {
    const user = users.find(u => u.username === username && u.password === password);
    if(user) {
      if(user.role === 'admin') {
        setSession(user);
        sessionStorage.setItem('ozn_session', JSON.stringify(user));
      } else {
        if (user.authorizedFirms && user.authorizedFirms.length === 1) {
          const finalSession = { ...user, firmaId: user.authorizedFirms[0] };
          setSession(finalSession);
          sessionStorage.setItem('ozn_session', JSON.stringify(finalSession));
        } else {
          setPendingUser(user);
        }
      }
    } else { alert("Hatalı kullanıcı kodu veya şifre!"); }
  };

  const handleFirmSelect = (firmId) => {
    if (!firmId) { setPendingUser(null); return; }
    const finalSession = { ...pendingUser, firmaId: firmId };
    setSession(finalSession);
    sessionStorage.setItem('ozn_session', JSON.stringify(finalSession));
    setPendingUser(null);
  };

  const handleLogout = () => {
    setSession(null); setPendingUser(null); sessionStorage.removeItem('ozn_session');
  };

  if (!session) return <LoginScreen onLogin={handleLogin} onFirmSelect={handleFirmSelect} pendingUser={pendingUser} firms={firms} />;
  if (session.role === 'admin') return <SaaSAdminApp session={session} onLogout={handleLogout} users={users} setUsers={setUsers} firms={firms} setFirms={setFirms} />;
  
  const firmDetails = firms.find(f => f.id === session.firmaId) || { id: session.firmaId, name: 'OZN TİCARET A.Ş.' };
  return <ERPClientApp session={session} onLogout={handleLogout} firmDetails={firmDetails} />;
}


// ==========================================
// 4. KURUMSAL MODALLAR VE FORMLAR
// ==========================================

const CariHesapModal = ({ onClose, onSave, initialData }) => {
  const [formData, setFormData] = useState(initialData || { 
    code: `120.01.${String(Math.floor(100 + Math.random()*899))}`, 
    desc: '', 
    type: 'Alıcı',
    taxOffice: '',
    taxNo: '',
    phone: '',
    city: 'İstanbul',
    riskLimit: 100000,
    balance: 0
  });
  return (
    <div className="modal-overlay">
      <div className="classic-window" style={{ width: '560px' }}>
        <div className="window-header"><div className="window-title">Cari Hesap Kartı Tanımı</div><X size={14} className="window-close" onClick={onClose} /></div>
        <div className="window-content">
          <div className="form-group"><label className="form-label">Cari Kodu:</label><input type="text" className="form-input" value={formData.code} onChange={e => setFormData({...formData, code: e.target.value})} autoFocus /></div>
          <div className="form-group"><label className="form-label">Ticari Unvanı:</label><input type="text" className="form-input" value={formData.desc} onChange={e => setFormData({...formData, desc: e.target.value.toUpperCase()})} placeholder="Örn: DEMİRSOY METAL SAN. TİC. LTD. ŞTİ." /></div>
          <div className="form-group"><label className="form-label">Cari Hesap Türü:</label>
            <select className="form-select" value={formData.type} onChange={e => setFormData({...formData, type: e.target.value})}>
              <option>Alıcı</option>
              <option>Satıcı</option>
              <option>Alıcı + Satıcı</option>
            </select>
          </div>
          <div className="form-row">
            <div className="form-col"><div className="form-group"><label className="form-label">Vergi Dairesi:</label><input type="text" className="form-input" value={formData.taxOffice} onChange={e => setFormData({...formData, taxOffice: e.target.value})} /></div></div>
            <div className="form-col"><div className="form-group"><label className="form-label">VKN / TCKN:</label><input type="text" className="form-input" value={formData.taxNo} onChange={e => setFormData({...formData, taxNo: e.target.value})} maxLength={11} /></div></div>
          </div>
          <div className="form-row">
            <div className="form-col"><div className="form-group"><label className="form-label">Telefon:</label><input type="text" className="form-input" value={formData.phone} onChange={e => setFormData({...formData, phone: e.target.value})} /></div></div>
            <div className="form-col"><div className="form-group"><label className="form-label">İl / Bölge:</label><input type="text" className="form-input" value={formData.city} onChange={e => setFormData({...formData, city: e.target.value})} /></div></div>
          </div>
          <div className="form-group"><label className="form-label">Risk Limiti (TL):</label><input type="number" className="form-input" value={formData.riskLimit} onChange={e => setFormData({...formData, riskLimit: e.target.value})} /></div>
        </div>
        <div className="window-footer">
          <button className="btn-classic primary" onClick={() => { if(!formData.desc) return alert("Unvan giriniz!"); onSave(formData); }}>Kaydet (F2)</button>
          <button className="btn-classic" onClick={onClose}>Vazgeç</button>
        </div>
      </div>
    </div>
  );
};

const StokModal = ({ onClose, onSave, initialData }) => {
  const [formData, setFormData] = useState(initialData || { 
    code: `STK.${String(Math.floor(100 + Math.random()*899))}`, 
    desc: '', 
    type: 'Ticari Mal', 
    unit: 'Adet',
    vat: 20,
    buyPrice: 0,
    sellPrice: 0,
    minStock: 5,
    balance: 0
  });
  return (
    <div className="modal-overlay">
      <div className="classic-window" style={{ width: '540px' }}>
        <div className="window-header"><div className="window-title">Malzeme (Stok) Kartı</div><X size={14} className="window-close" onClick={onClose} /></div>
        <div className="window-content">
          <div className="form-group"><label className="form-label">Malzeme Kodu:</label><input type="text" className="form-input" value={formData.code} onChange={e => setFormData({...formData, code: e.target.value})} autoFocus /></div>
          <div className="form-group"><label className="form-label">Malzeme Adı:</label><input type="text" className="form-input" value={formData.desc} onChange={e => setFormData({...formData, desc: e.target.value.toUpperCase()})} /></div>
          <div className="form-row">
            <div className="form-col">
              <div className="form-group"><label className="form-label">Kart Türü:</label>
                <select className="form-select" value={formData.type} onChange={e => setFormData({...formData, type: e.target.value})}>
                  <option>Ticari Mal</option><option>Hammadde</option><option>Yarı Mamul</option><option>Mamul</option><option>Sarf Malzemesi</option>
                </select>
              </div>
            </div>
            <div className="form-col">
              <div className="form-group"><label className="form-label">Ana Birim:</label>
                <select className="form-select" value={formData.unit} onChange={e => setFormData({...formData, unit: e.target.value})}>
                  <option>Adet</option><option>Kg</option><option>Koli</option><option>Kutu</option><option>Metre</option><option>Litre</option>
                </select>
              </div>
            </div>
          </div>
          <div className="form-row">
            <div className="form-col"><div className="form-group"><label className="form-label">Alış Fiyatı (TL):</label><input type="number" className="form-input" value={formData.buyPrice} onChange={e => setFormData({...formData, buyPrice: e.target.value})} /></div></div>
            <div className="form-col"><div className="form-group"><label className="form-label">Satış Fiyatı (TL):</label><input type="number" className="form-input" value={formData.sellPrice} onChange={e => setFormData({...formData, sellPrice: e.target.value})} /></div></div>
          </div>
          <div className="form-row">
            <div className="form-col">
              <div className="form-group"><label className="form-label">KDV Oranı (%):</label>
                <select className="form-select" value={formData.vat} onChange={e => setFormData({...formData, vat: parseInt(e.target.value)})}>
                  <option value={20}>%20</option><option value={10}>%10</option><option value={1}>%1</option><option value={0}>%0</option>
                </select>
              </div>
            </div>
            <div className="form-col"><div className="form-group"><label className="form-label">Asgari Stok:</label><input type="number" className="form-input" value={formData.minStock} onChange={e => setFormData({...formData, minStock: e.target.value})} /></div></div>
          </div>
        </div>
        <div className="window-footer"><button className="btn-classic primary" onClick={() => { if(!formData.desc) return alert("Açıklama giriniz!"); onSave(formData); }}>Kaydet (F2)</button><button className="btn-classic" onClick={onClose}>Vazgeç</button></div>
      </div>
    </div>
  );
};

const MalzemeFisiModal = ({ onClose, stoklar, onSave }) => {
  const [formData, setFormData] = useState({ code: `MLZ-${Math.floor(1000+Math.random()*9000)}`, date: new Date().toISOString().split('T')[0], type: 'Sayım Fazlası (Giriş)', stokId: '', qty: 1, desc: '' });
  return (
    <div className="modal-overlay">
      <div className="classic-window" style={{ width: '500px' }}>
        <div className="window-header"><div className="window-title">Malzeme Yönetim Fişi</div><X size={14} className="window-close" onClick={onClose} /></div>
        <div className="window-content">
          <div className="form-group"><label className="form-label">Fiş No / Tarih:</label><input type="text" className="form-input" value={formData.code} readOnly style={{marginRight: 8}} /><input type="date" className="form-input" value={formData.date} onChange={e => setFormData({...formData, date: e.target.value})} /></div>
          <div className="form-group"><label className="form-label">Fiş Türü:</label>
            <select className="form-select" value={formData.type} onChange={e => setFormData({...formData, type: e.target.value})}>
              <option>Devir Fişi (Giriş)</option>
              <option>Sayım Fazlası (Giriş)</option>
              <option>Üretimden Giriş Fişi</option>
              <option>Sarf Fişi (Çıkış)</option>
              <option>Fire Fişi (Çıkış)</option>
              <option>Sayım Eksiği (Çıkış)</option>
            </select>
          </div>
          <div className="form-group"><label className="form-label">Malzeme:</label>
            <select className="form-select" value={formData.stokId} onChange={e => setFormData({...formData, stokId: e.target.value})}>
              <option value="">-- Malzeme Seçin --</option>
              {stoklar.map(s => <option key={s.id} value={s.id}>{s.code} - {s.desc} (Mevcut: {s.balance})</option>)}
            </select>
          </div>
          <div className="form-group"><label className="form-label">Miktar:</label><input type="number" className="form-input" value={formData.qty} onChange={e => setFormData({...formData, qty: e.target.value})} /></div>
          <div className="form-group"><label className="form-label">Açıklama:</label><input type="text" className="form-input" value={formData.desc} onChange={e => setFormData({...formData, desc: e.target.value})} /></div>
        </div>
        <div className="window-footer"><button className="btn-classic primary" onClick={() => { if(!formData.stokId) return alert("Malzeme seçin!"); onSave(formData); }}>Kaydet (F2)</button><button className="btn-classic" onClick={onClose}>Vazgeç</button></div>
      </div>
    </div>
  );
};

const BankaKasaModal = ({ onClose, onSave, type, initialData }) => {
  const [formData, setFormData] = useState(initialData || { code: `${type.toUpperCase()}.0${Math.floor(1+Math.random()*9)}`, desc: '', branch: '', iban: '', currency: 'TRY', balance: 0 });
  return (
    <div className="modal-overlay">
      <div className="classic-window" style={{ width: '500px' }}>
        <div className="window-header"><div className="window-title">{type} Kartı Tanımı</div><X size={14} className="window-close" onClick={onClose} /></div>
        <div className="window-content">
          <div className="form-group"><label className="form-label">{type} Kodu:</label><input type="text" className="form-input" value={formData.code} onChange={e => setFormData({...formData, code: e.target.value})} autoFocus /></div>
          <div className="form-group"><label className="form-label">Hesap Adı:</label><input type="text" className="form-input" value={formData.desc} onChange={e => setFormData({...formData, desc: e.target.value.toUpperCase()})} /></div>
          {type === 'Banka' && (
            <>
              <div className="form-group"><label className="form-label">Şube Adı:</label><input type="text" className="form-input" value={formData.branch} onChange={e => setFormData({...formData, branch: e.target.value})} /></div>
              <div className="form-group"><label className="form-label">IBAN No:</label><input type="text" className="form-input" value={formData.iban} onChange={e => setFormData({...formData, iban: e.target.value})} placeholder="TR00 0000 ..." /></div>
            </>
          )}
          <div className="form-group"><label className="form-label">Para Birimi:</label>
            <select className="form-select" value={formData.currency} onChange={e => setFormData({...formData, currency: e.target.value})}>
              <option>TRY</option><option>USD</option><option>EUR</option>
            </select>
          </div>
        </div>
        <div className="window-footer"><button className="btn-classic primary" onClick={() => onSave(formData)}>Kaydet (F2)</button><button className="btn-classic" onClick={onClose}>Vazgeç</button></div>
      </div>
    </div>
  );
};

const FaturaModal = ({ onClose, cariler, stoklar, faturaTipi, isOrder, onSave }) => {
  const prefix = isOrder ? 'SIP' : (faturaTipi.includes('Satınalma') ? 'ALF' : 'STF');
  const [formData, setFormData] = useState({ code: `${prefix}-${Math.floor(10000+Math.random()*90000)}`, date: new Date().toISOString().split('T')[0], cariId: '', docNo: '' });
  const [lines, setLines] = useState([{ id: 1, stokId: '', qty: 1, price: 0, discount: 0, vat: 20, total: 0 }]);

  const updateLine = (id, field, val) => {
    setLines(lines.map(l => {
      if (l.id !== id) return l;
      const updated = { ...l, [field]: val };
      if (field === 'stokId') {
        const st = stoklar.find(s => s.id.toString() === val.toString());
        if (st) {
          updated.price = faturaTipi.includes('Satınalma') ? (st.buyPrice || 0) : (st.sellPrice || 0);
          updated.vat = st.vat || 20;
        }
      }
      const raw = (parseFloat(updated.qty) || 0) * (parseFloat(updated.price) || 0);
      const disc = raw * ((parseFloat(updated.discount) || 0) / 100);
      const net = raw - disc;
      const withVat = net * (1 + ((parseFloat(updated.vat) || 0) / 100));
      updated.total = withVat;
      return updated;
    }));
  };

  const addLine = () => setLines([...lines, { id: Date.now(), stokId: '', qty: 1, price: 0, discount: 0, vat: 20, total: 0 }]);
  const removeLine = (id) => { if (lines.length > 1) setLines(lines.filter(l => l.id !== id)); };

  const subTotal = lines.reduce((acc, l) => acc + ((parseFloat(l.qty)||0) * (parseFloat(l.price)||0)), 0);
  const discountTotal = lines.reduce((acc, l) => acc + ((parseFloat(l.qty)||0) * (parseFloat(l.price)||0) * ((parseFloat(l.discount)||0)/100)), 0);
  const vatTotal = lines.reduce((acc, l) => {
    const net = ((parseFloat(l.qty)||0) * (parseFloat(l.price)||0)) * (1 - ((parseFloat(l.discount)||0)/100));
    return acc + (net * ((parseFloat(l.vat)||0)/100));
  }, 0);
  const netTotal = subTotal - discountTotal + vatTotal;

  const handleSave = () => {
    if (!formData.cariId) return alert("Lütfen Cari Hesap seçiniz!");
    onSave({ 
      ...formData, 
      type: faturaTipi, 
      subTotal: subTotal - discountTotal, 
      discountTotal, 
      vatTotal, 
      total: netTotal.toFixed(2), 
      lines 
    });
  };

  return (
    <div className="modal-overlay">
      <div className="classic-window" style={{ width: '920px', height: '580px' }}>
        <div className="window-header"><div className="window-title">{faturaTipi} (Detaylı Kalem Girişi)</div><X size={14} className="window-close" onClick={onClose} /></div>
        <div className="window-content" style={{ padding: 8 }}>
          <div className="master-detail-top">
            <div className="form-row">
              <div className="form-col">
                <div className="form-group"><label className="form-label">Fiş / Fatura No:</label><input type="text" className="form-input" value={formData.code} onChange={e=>setFormData({...formData, code: e.target.value})} /></div>
                <div className="form-group"><label className="form-label">Cari Hesap:</label>
                  <select className="form-select" value={formData.cariId} onChange={e => setFormData({...formData, cariId: e.target.value})}>
                    <option value="">-- Cari Hesap Seçiniz --</option>
                    {cariler.map(c => <option key={c.id} value={c.id}>{c.code} - {c.desc} (Bakiye: {formatMoney(c.balance)} TL)</option>)}
                  </select>
                </div>
              </div>
              <div className="form-col">
                <div className="form-group"><label className="form-label">Belge Tarihi:</label><input type="date" className="form-input" value={formData.date} onChange={e => setFormData({...formData, date: e.target.value})} /></div>
                <div className="form-group"><label className="form-label">Belge No / İrsaliye:</label><input type="text" className="form-input" value={formData.docNo} onChange={e => setFormData({...formData, docNo: e.target.value})} placeholder="E-Arşiv / E-Fatura No" /></div>
              </div>
            </div>
          </div>

          <div style={{marginBottom: 6, display: 'flex', justifyContent: 'space-between', alignItems: 'center'}}>
            <span style={{fontWeight: 700, color: '#0f2c59'}}>Malzeme / Hizmet Satırları</span>
            <button className="btn-classic" onClick={addLine}><Plus size={12} /> Yeni Satır Ekle</button>
          </div>

          <div className="master-detail-grid">
            <table className="data-grid">
              <thead>
                <tr>
                  <th style={{width: 30}}>#</th>
                  <th>Malzeme Kodu / Açıklaması</th>
                  <th style={{width: 80, textAlign: 'right'}}>Miktar</th>
                  <th style={{width: 110, textAlign: 'right'}}>Birim Fiyat</th>
                  <th style={{width: 75, textAlign: 'right'}}>İsk. %</th>
                  <th style={{width: 75, textAlign: 'center'}}>KDV %</th>
                  <th style={{width: 125, textAlign: 'right'}}>Satır Net (KDV Dahil)</th>
                  <th style={{width: 40}}></th>
                </tr>
              </thead>
              <tbody>
                {lines.map((l, idx) => (
                  <tr key={l.id}>
                    <td style={{textAlign: 'center'}}>{idx + 1}</td>
                    <td>
                      <select style={{width: '100%', border: 'none', background: 'transparent', fontSize: 12}} value={l.stokId} onChange={e => updateLine(l.id, 'stokId', e.target.value)}>
                        <option value="">-- Malzeme Seçin --</option>
                        {stoklar.map(s => <option key={s.id} value={s.id}>{s.code} - {s.desc} (Stok: {s.balance})</option>)}
                      </select>
                    </td>
                    <td><input type="number" value={l.qty} onChange={e => updateLine(l.id, 'qty', e.target.value)} style={{width: '100%', border: '1px solid #ddd', textAlign: 'right', padding: 2}} /></td>
                    <td><input type="number" value={l.price} onChange={e => updateLine(l.id, 'price', e.target.value)} style={{width: '100%', border: '1px solid #ddd', textAlign: 'right', padding: 2}} /></td>
                    <td><input type="number" value={l.discount} onChange={e => updateLine(l.id, 'discount', e.target.value)} style={{width: '100%', border: '1px solid #ddd', textAlign: 'right', padding: 2}} /></td>
                    <td>
                      <select value={l.vat} onChange={e => updateLine(l.id, 'vat', e.target.value)} style={{width: '100%', border: 'none', textAlign: 'center'}}>
                        <option value={20}>%20</option><option value={10}>%10</option><option value={1}>%1</option><option value={0}>%0</option>
                      </select>
                    </td>
                    <td style={{textAlign: 'right', fontWeight: 700}}>{formatMoney(l.total)} TL</td>
                    <td style={{textAlign: 'center'}}><Trash2 size={13} color="#c92a2a" style={{cursor: 'pointer'}} onClick={() => removeLine(l.id)} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="master-detail-bottom">
            <div style={{fontSize: 11, color: '#57606f'}}>
              * Kaydedildiğinde Stok, Cari Bakiye ve Genel Muhasebe (120/153/191/320/391/600) hesaplarına otomatik entegre edilir.
            </div>
            <div style={{display: 'flex', gap: 18, fontSize: 12}}>
              <div>Ara Toplam: <b>{formatMoney(subTotal)} TL</b></div>
              <div>İskonto: <b>{formatMoney(discountTotal)} TL</b></div>
              <div>Toplam KDV: <b>{formatMoney(vatTotal)} TL</b></div>
              <div style={{color: '#0f2c59', fontSize: 14}}>Genel Toplam: <span>{formatMoney(netTotal)} TL</span></div>
            </div>
          </div>
        </div>
        <div className="window-footer">
          <button className="btn-classic primary" onClick={handleSave}>Kaydet ve Muhasebeleştir (F2)</button>
          <button className="btn-classic" onClick={onClose}>Vazgeç</button>
        </div>
      </div>
    </div>
  );
};

const FinansFisiModal = ({ onClose, cariler, aracListesi, arac, onSave }) => {
  const [formData, setFormData] = useState({ 
    code: `${arac === 'Kasa' ? 'KAS' : 'BNK'}-${Math.floor(1000+Math.random()*9000)}`, 
    date: new Date().toISOString().split('T')[0],
    type: arac === 'Kasa' ? 'Tahsilat Fişi' : 'Gelen Havale / Tahsilat',
    aracId: aracListesi[0]?.id || '', 
    cariId: '', 
    total: 0,
    desc: ''
  });
  const handleSave = () => { 
    if (!formData.aracId || !formData.cariId || !formData.total) return alert("Lütfen Kasa/Banka, Cari ve Tutar bilgilerini eksiksiz giriniz!"); 
    onSave({ ...formData, arac }); 
  };
  return (
    <div className="modal-overlay">
      <div className="classic-window" style={{ width: '520px' }}>
        <div className="window-header"><div className="window-title">{arac} İşlem Fişi</div><X size={14} className="window-close" onClick={onClose} /></div>
        <div className="window-content">
          <div className="form-group"><label className="form-label">Fiş No / Tarih:</label><input type="text" className="form-input" value={formData.code} readOnly style={{marginRight: 8}} /><input type="date" className="form-input" value={formData.date} onChange={e => setFormData({...formData, date: e.target.value})} /></div>
          <div className="form-group"><label className="form-label">İşlem Türü:</label>
            <select className="form-select" value={formData.type} onChange={e => setFormData({...formData, type: e.target.value})}>
              {arac === 'Kasa' ? (
                <>
                  <option>Tahsilat Fişi (Para Girişi)</option>
                  <option>Tediye (Ödeme) Fişi (Para Çıkışı)</option>
                </>
              ) : (
                <>
                  <option>Gelen Havale / Tahsilat</option>
                  <option>Gönderilen Havale / EFT (Ödeme)</option>
                </>
              )}
            </select>
          </div>
          <div className="form-group"><label className="form-label">İşlem Yapılan {arac}:</label>
            <select className="form-select" value={formData.aracId} onChange={e => setFormData({...formData, aracId: e.target.value})}>
              <option value="">-- {arac} Seçin --</option>
              {aracListesi.map(a => <option key={a.id} value={a.id}>{a.code} - {a.desc} ({formatMoney(a.balance)} TL)</option>)}
            </select>
          </div>
          <div className="form-group"><label className="form-label">Karşı Cari Hesap:</label>
            <select className="form-select" value={formData.cariId} onChange={e => setFormData({...formData, cariId: e.target.value})}>
              <option value="">-- Cari Seçin --</option>
              {cariler.map(c => <option key={c.id} value={c.id}>{c.code} - {c.desc} ({formatMoney(c.balance)} TL)</option>)}
            </select>
          </div>
          <div className="form-group"><label className="form-label">İşlem Tutarı (TL):</label><input type="number" className="form-input" value={formData.total} onChange={e => setFormData({...formData, total: e.target.value})} /></div>
          <div className="form-group"><label className="form-label">Açıklama:</label><input type="text" className="form-input" value={formData.desc} onChange={e => setFormData({...formData, desc: e.target.value})} placeholder="Tahsilat / Ödeme açıklaması..." /></div>
        </div>
        <div className="window-footer"><button className="btn-classic primary" onClick={handleSave}>Kaydet ve Muhasebeleştir (F2)</button><button className="btn-classic" onClick={onClose}>Vazgeç</button></div>
      </div>
    </div>
  );
};

const CariFisiModal = ({ onClose, cariler, onSave }) => {
  const [formData, setFormData] = useState({ code: `CFI-${Math.floor(1000+Math.random()*9000)}`, date: new Date().toISOString().split('T')[0], cariId: '', type: 'Borç Dekontu', total: 0, desc: '' });
  const handleSave = () => { if (!formData.cariId || !formData.total) return alert("Cari ve Tutar giriniz!"); onSave({ ...formData }); };
  return (
    <div className="modal-overlay">
      <div className="classic-window" style={{ width: '500px' }}>
        <div className="window-header"><div className="window-title">Cari Hesap Fişi (Dekont / Açılış)</div><X size={14} className="window-close" onClick={onClose} /></div>
        <div className="window-content">
          <div className="form-group"><label className="form-label">Fiş Numarası:</label><input type="text" className="form-input" value={formData.code} readOnly /></div>
          <div className="form-group"><label className="form-label">Fiş Türü:</label>
            <select className="form-select" value={formData.type} onChange={e => setFormData({...formData, type: e.target.value})}>
              <option>Borç Dekontu</option>
              <option>Alacak Dekontu</option>
              <option>Açılış Fişi (Borç)</option>
            </select>
          </div>
          <div className="form-group"><label className="form-label">Cari Hesap:</label>
            <select className="form-select" value={formData.cariId} onChange={e => setFormData({...formData, cariId: e.target.value})}>
              <option value="">-- Cari Seçin --</option>
              {cariler.map(c => <option key={c.id} value={c.id}>{c.code} - {c.desc}</option>)}
            </select>
          </div>
          <div className="form-group"><label className="form-label">Tutar (TL):</label><input type="number" className="form-input" value={formData.total} onChange={e => setFormData({...formData, total: e.target.value})} /></div>
          <div className="form-group"><label className="form-label">Açıklama:</label><input type="text" className="form-input" value={formData.desc} onChange={e => setFormData({...formData, desc: e.target.value})} /></div>
        </div>
        <div className="window-footer"><button className="btn-classic primary" onClick={handleSave}>Kaydet (F2)</button><button className="btn-classic" onClick={onClose}>Vazgeç</button></div>
      </div>
    </div>
  );
};

const CekSenetModal = ({ onClose, cariler, onSave }) => {
  const [formData, setFormData] = useState({
    code: `BRD-${Math.floor(1000+Math.random()*9000)}`,
    vade: new Date(Date.now() + 30*86400000).toISOString().split('T')[0],
    type: 'Müşteri Çeki (Alınan)',
    cariId: '',
    bankName: 'Garanti BBVA',
    checkNo: String(Math.floor(1000000+Math.random()*9000000)),
    total: 0
  });
  return (
    <div className="modal-overlay">
      <div className="classic-window" style={{ width: '520px' }}>
        <div className="window-header"><div className="window-title">Çek / Senet Giriş-Çıkış Bordrosu</div><X size={14} className="window-close" onClick={onClose} /></div>
        <div className="window-content">
          <div className="form-group"><label className="form-label">Bordro No:</label><input type="text" className="form-input" value={formData.code} readOnly /></div>
          <div className="form-group"><label className="form-label">Bordro Türü:</label>
            <select className="form-select" value={formData.type} onChange={e => setFormData({...formData, type: e.target.value})}>
              <option>Müşteri Çeki (Alınan)</option>
              <option>Müşteri Senedi (Alınan)</option>
              <option>Kendi Çekimiz (Verilen)</option>
              <option>Borç Senedimiz (Verilen)</option>
            </select>
          </div>
          <div className="form-group"><label className="form-label">Cari Hesap:</label>
            <select className="form-select" value={formData.cariId} onChange={e => setFormData({...formData, cariId: e.target.value})}>
              <option value="">-- Cari Hesap Seçin --</option>
              {cariler.map(c => <option key={c.id} value={c.id}>{c.code} - {c.desc}</option>)}
            </select>
          </div>
          <div className="form-group"><label className="form-label">Vade Tarihi:</label><input type="date" className="form-input" value={formData.vade} onChange={e => setFormData({...formData, vade: e.target.value})} /></div>
          <div className="form-row">
            <div className="form-col"><div className="form-group"><label className="form-label">Banka Adı:</label><input type="text" className="form-input" value={formData.bankName} onChange={e => setFormData({...formData, bankName: e.target.value})} /></div></div>
            <div className="form-col"><div className="form-group"><label className="form-label">Çek/Senet No:</label><input type="text" className="form-input" value={formData.checkNo} onChange={e => setFormData({...formData, checkNo: e.target.value})} /></div></div>
          </div>
          <div className="form-group"><label className="form-label">Tutar (TL):</label><input type="number" className="form-input" value={formData.total} onChange={e => setFormData({...formData, total: e.target.value})} /></div>
        </div>
        <div className="window-footer">
          <button className="btn-classic primary" onClick={() => { if(!formData.cariId || !formData.total) return alert("Cari ve Tutar giriniz!"); onSave(formData); }}>Kaydet ve Muhasebeleştir</button>
          <button className="btn-classic" onClick={onClose}>Vazgeç</button>
        </div>
      </div>
    </div>
  );
};

const HesapEkleModal = ({ onClose, onSave }) => {
  const [formData, setFormData] = useState({ code: '', desc: '', type: 'Aktif' });
  return (
    <div className="modal-overlay">
      <div className="classic-window" style={{ width: '460px' }}>
        <div className="window-header"><div className="window-title">Yeni Muhasebe Hesap Kartı</div><X size={14} className="window-close" onClick={onClose} /></div>
        <div className="window-content">
          <div className="form-group"><label className="form-label">Hesap Kodu:</label><input type="text" className="form-input" value={formData.code} onChange={e => setFormData({...formData, code: e.target.value})} placeholder="Örn: 100.01" autoFocus /></div>
          <div className="form-group"><label className="form-label">Hesap Adı:</label><input type="text" className="form-input" value={formData.desc} onChange={e => setFormData({...formData, desc: e.target.value.toUpperCase()})} placeholder="Örn: MERKEZ TL KASA" /></div>
          <div className="form-group"><label className="form-label">Hesap Karakteri:</label>
            <select className="form-select" value={formData.type} onChange={e => setFormData({...formData, type: e.target.value})}>
              <option>Aktif</option><option>Pasif</option><option>Gelir</option><option>Gider</option>
            </select>
          </div>
        </div>
        <div className="window-footer"><button className="btn-classic primary" onClick={() => { if(!formData.code || !formData.desc) return alert("Eksik bilgi!"); onSave(formData); }}>Kaydet</button><button className="btn-classic" onClick={onClose}>Vazgeç</button></div>
      </div>
    </div>
  );
};

const MuhasebeFisiModal = ({ onClose, hesapPlani, onSave }) => {
  const [formData, setFormData] = useState({ type: 'Mahsup Fişi', desc: '' });
  const [lines, setLines] = useState([
    { id: 1, accountCode: '770', borc: 0, alacak: 0, desc: '' },
    { id: 2, accountCode: '100', borc: 0, alacak: 0, desc: '' }
  ]);
  const totalBorc = lines.reduce((s, l) => s + (parseFloat(l.borc) || 0), 0);
  const totalAlacak = lines.reduce((s, l) => s + (parseFloat(l.alacak) || 0), 0);

  const updateLine = (id, f, v) => setLines(lines.map(l => l.id === id ? { ...l, [f]: v } : l));
  const addLine = () => setLines([...lines, { id: Date.now(), accountCode: '100', borc: 0, alacak: 0, desc: '' }]);

  const handleSave = () => {
    if (Math.abs(totalBorc - totalAlacak) > 0.01 || totalBorc === 0) {
      return alert("Muhasebe denkliği gereği Toplam Borç ve Toplam Alacak birbirine eşit ve sıfırdan büyük olmalıdır!");
    }
    onSave({ ...formData, lines });
  };

  return (
    <div className="modal-overlay">
      <div className="classic-window" style={{ width: '760px', height: '480px' }}>
        <div className="window-header"><div className="window-title">Muhasebe Yevmiye / Mahsup Fişi</div><X size={14} className="window-close" onClick={onClose} /></div>
        <div className="window-content">
          <div className="master-detail-top">
            <div className="form-row">
              <div className="form-col"><div className="form-group"><label className="form-label">Fiş Türü:</label><select className="form-select" value={formData.type} onChange={e=>setFormData({...formData, type: e.target.value})}><option>Mahsup Fişi</option><option>Tahsil Fişi</option><option>Tediye Fişi</option><option>Açılış Fişi</option></select></div></div>
              <div className="form-col"><div className="form-group"><label className="form-label">Genel Açıklama:</label><input type="text" className="form-input" value={formData.desc} onChange={e=>setFormData({...formData, desc: e.target.value})} placeholder="Yevmiye açıklaması..." /></div></div>
            </div>
          </div>
          <div style={{marginBottom: 6, textAlign: 'right'}}><button className="btn-classic" onClick={addLine}><Plus size={12}/> Satır Ekle</button></div>
          <div className="master-detail-grid">
            <table className="data-grid">
              <thead><tr><th>Muhasebe Hesabı</th><th>Satır Açıklaması</th><th style={{width: 120, textAlign: 'right'}}>Borç (TL)</th><th style={{width: 120, textAlign: 'right'}}>Alacak (TL)</th></tr></thead>
              <tbody>
                {lines.map(l => (
                  <tr key={l.id}>
                    <td>
                      <select style={{width: '100%', border: 'none'}} value={l.accountCode} onChange={e => updateLine(l.id, 'accountCode', e.target.value)}>
                        {hesapPlani.map(h => <option key={h.code} value={h.code}>{h.code} - {h.desc}</option>)}
                      </select>
                    </td>
                    <td><input type="text" value={l.desc} onChange={e => updateLine(l.id, 'desc', e.target.value)} style={{width: '100%', border: '1px solid #ddd', padding: 2}} /></td>
                    <td><input type="number" value={l.borc} onChange={e => updateLine(l.id, 'borc', e.target.value)} style={{width: '100%', border: '1px solid #ddd', textAlign: 'right', padding: 2}} /></td>
                    <td><input type="number" value={l.alacak} onChange={e => updateLine(l.id, 'alacak', e.target.value)} style={{width: '100%', border: '1px solid #ddd', textAlign: 'right', padding: 2}} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="master-detail-bottom">
            <div style={{color: Math.abs(totalBorc - totalAlacak) < 0.01 ? '#2b8a3e' : '#c92a2a'}}>
              Fark: {formatMoney(Math.abs(totalBorc - totalAlacak))} TL {Math.abs(totalBorc - totalAlacak) < 0.01 ? '(Bakiye Denk)' : '(Borç/Alacak Eşit Değil!)'}
            </div>
            <div>Toplam Borç: <b>{formatMoney(totalBorc)} TL</b> | Toplam Alacak: <b>{formatMoney(totalAlacak)} TL</b></div>
          </div>
        </div>
        <div className="window-footer"><button className="btn-classic primary" onClick={handleSave}>Yevmiye Kaydını Onayla</button><button className="btn-classic" onClick={onClose}>Vazgeç</button></div>
      </div>
    </div>
  );
};

const EkstreModal = ({ target, onClose, faturalar, finansFisleri, cariFisleri, cekSenetler, malzemeFisleri }) => {
  const isCari = target.type === 'cari';
  const item = target.item;

  let movements = [];
  if (isCari) {
    faturalar.filter(f => f.cariId.toString() === item.id.toString()).forEach(f => {
      const isAlim = f.type.includes('Satınalma');
      movements.push({ date: f.date, code: f.code, type: f.type, borc: isAlim ? 0 : parseFloat(f.total), alacak: isAlim ? parseFloat(f.total) : 0 });
    });
    finansFisleri.filter(f => f.cariId.toString() === item.id.toString()).forEach(f => {
      const isTahsilat = f.type.includes('Tahsilat') || f.type.includes('Gelen');
      movements.push({ date: f.date, code: f.code, type: f.type, borc: isTahsilat ? 0 : parseFloat(f.total), alacak: isTahsilat ? parseFloat(f.total) : 0 });
    });
    cariFisleri.filter(f => f.cariId.toString() === item.id.toString()).forEach(f => {
      const isAlacak = f.type === 'Alacak Dekontu';
      movements.push({ date: f.date, code: f.code, type: f.type, borc: isAlacak ? 0 : parseFloat(f.total), alacak: isAlacak ? parseFloat(f.total) : 0 });
    });
    cekSenetler.filter(c => c.cariId.toString() === item.id.toString()).forEach(c => {
      const isAlinan = c.type.includes('Müşteri');
      movements.push({ date: c.vade, code: c.code, type: c.type, borc: isAlinan ? 0 : parseFloat(c.total), alacak: isAlinan ? parseFloat(c.total) : 0 });
    });
  } else {
    faturalar.forEach(f => {
      f.lines?.filter(l => l.stokId.toString() === item.id.toString()).forEach(l => {
        const isAlim = f.type.includes('Satınalma');
        movements.push({ date: f.date, code: f.code, type: f.type, giris: isAlim ? l.qty : 0, cikis: isAlim ? 0 : l.qty, price: l.price });
      });
    });
    malzemeFisleri.filter(m => m.stokId.toString() === item.id.toString()).forEach(m => {
      const isGiris = m.type.includes('Giriş') || m.type.includes('Fazlası') || m.type.includes('Devir');
      movements.push({ date: m.date, code: m.code, type: m.type, giris: isGiris ? m.qty : 0, cikis: isGiris ? 0 : m.qty, price: item.buyPrice || 0 });
    });
  }

  return (
    <div className="modal-overlay">
      <div className="classic-window" style={{ width: '780px', height: '480px' }}>
        <div className="window-header">
          <div className="window-title">{isCari ? 'Cari Hesap Ekstresi (Hareket Dökümü)' : 'Malzeme (Stok) Hareket Ekstresi'} - [{item.code} {item.desc}]</div>
          <X size={14} className="window-close" onClick={onClose} />
        </div>
        <div className="window-content">
          <div className="master-detail-grid">
            <table className="data-grid">
              <thead>
                {isCari ? (
                  <tr><th>Tarih</th><th>Fiş / Belge No</th><th>İşlem Türü</th><th style={{textAlign: 'right'}}>Borç (TL)</th><th style={{textAlign: 'right'}}>Alacak (TL)</th></tr>
                ) : (
                  <tr><th>Tarih</th><th>Belge No</th><th>Hareket Türü</th><th style={{textAlign: 'right'}}>Giren Miktar</th><th style={{textAlign: 'right'}}>Çıkan Miktar</th><th style={{textAlign: 'right'}}>Birim Fiyat</th></tr>
                )}
              </thead>
              <tbody>
                {movements.length === 0 ? (
                  <tr><td colSpan={6} style={{textAlign: 'center', padding: 20}}>Bu karta ait hareket kaydı bulunamadı.</td></tr>
                ) : movements.map((m, i) => (
                  <tr key={i}>
                    <td>{m.date}</td>
                    <td><b>{m.code}</b></td>
                    <td>{m.type}</td>
                    {isCari ? (
                      <>
                        <td style={{textAlign: 'right', color: '#2b8a3e', fontWeight: 600}}>{m.borc > 0 ? formatMoney(m.borc) : '-'}</td>
                        <td style={{textAlign: 'right', color: '#c92a2a', fontWeight: 600}}>{m.alacak > 0 ? formatMoney(m.alacak) : '-'}</td>
                      </>
                    ) : (
                      <>
                        <td style={{textAlign: 'right', color: '#2b8a3e', fontWeight: 600}}>{m.giris > 0 ? m.giris : '-'}</td>
                        <td style={{textAlign: 'right', color: '#c92a2a', fontWeight: 600}}>{m.cikis > 0 ? m.cikis : '-'}</td>
                        <td style={{textAlign: 'right'}}>{formatMoney(m.price)} TL</td>
                      </>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="master-detail-bottom">
            <div>Kart Kodu: <b>{item.code}</b></div>
            <div>Güncel Net Bakiye: <b>{isCari ? `${formatMoney(item.balance)} TL` : `${item.balance} ${item.unit}`}</b></div>
          </div>
        </div>
        <div className="window-footer">
          <button className="btn-classic" onClick={() => window.print()}><Printer size={13}/> Ekstreyi Yazdır</button>
          <button className="btn-classic primary" onClick={onClose}>Kapat</button>
        </div>
      </div>
    </div>
  );
};

export default App;
