import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileSpreadsheet, 
  Upload, 
  Download, 
  Plus, 
  Trash2, 
  Edit3, 
  Search, 
  LogOut, 
  Phone, 
  Lock, 
  AlertCircle,
  Image as ImageIcon,
  ChevronLeft,
  ChevronRight,
  X,
  Check,
  Sparkles,
  Eye,
  ExternalLink,
  UploadCloud,
  Loader2,
  Maximize2,
  ArrowLeft,
  RefreshCw
} from 'lucide-react';

// ==========================================
// 🔐 CLOUDINARY .ENV CONFIGURATION
// ==========================================
const CLOUDINARY_CLOUD_NAME = import.meta.env.VITE_CLOUDINARY_CLOUD_NAME || 'mq6c2sfk';
const CLOUDINARY_API_KEY = import.meta.env.VITE_CLOUDINARY_API_KEY || '';
const CLOUDINARY_API_SECRET = import.meta.env.VITE_CLOUDINARY_API_SECRET || '';
const CLOUDINARY_UPLOAD_PRESET = import.meta.env.VITE_CLOUDINARY_UPLOAD_PRESET || '';

// SHA-1 Signature Generator
async function generateSHA1Signature(message) {
  const msgBuffer = new TextEncoder().encode(message);
  const hashBuffer = await crypto.subtle.digest('SHA-1', msgBuffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

// Exact 11 Columns
const COLUMNS = [
  'Image',
  'Name',
  'Price',
  'Original Price',
  'Quantity',
  'Sub-Category',
  'Category',
  'Hindi Name',
  'Hinglish Name',
  'Indian Category',
  'Indian Sub-Category'
];

// IndexedDB Helper for high-capacity 33,000+ dataset storage
const DB_NAME = 'CatalogStoreDB';
const STORE_NAME = 'MasterCatalogStore';

function openDatabase() {
  return new Promise((resolve) => {
    if (!window.indexedDB) {
      resolve(null);
      return;
    }
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => resolve(null);
  });
}

async function saveToDB(key, val) {
  try {
    const db = await openDatabase();
    if (!db) return;
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).put(val, key);
  } catch (err) {
    console.warn(err);
  }
}

async function loadFromDB(key) {
  try {
    const db = await openDatabase();
    if (!db) return null;
    return new Promise((resolve) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const req = tx.objectStore(STORE_NAME).get(key);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  } catch {
    return null;
  }
}

// Hinglish to Hindi Transliteration Helper
async function convertHinglishToHindi(text) {
  if (!text || !text.trim()) return '';
  try {
    const res = await fetch(
      `https://inputtools.google.com/request?text=${encodeURIComponent(text)}&itc=hi-t-i0-und&num=1`
    );
    const data = await res.json();
    if (data && data[0] === 'SUCCESS' && data[1]) {
      return data[1].map((item) => item[1][0]).join(' ');
    }
  } catch (e) {
    console.warn('Transliteration fallback:', e);
  }
  return text;
}

export default function App() {
  // ------------------------------------
  // CHECK FOR INDIVIDUAL PRODUCT TAB URL
  // ------------------------------------
  const [individualProduct, setIndividualProduct] = useState(() => {
    try {
      const urlParams = new URLSearchParams(window.location.search);
      const isProductView = urlParams.get('product');
      if (isProductView) {
        const saved = sessionStorage.getItem('individual_view_item');
        return saved ? JSON.parse(saved) : null;
      }
    } catch {}
    return null;
  });

  const handleOpenIndividualTab = (item) => {
    try {
      sessionStorage.setItem('individual_view_item', JSON.stringify(item));
      const url = `${window.location.origin}${window.location.pathname}?product=${encodeURIComponent(item['Name'])}`;
      window.open(url, '_blank');
    } catch {
      window.open(item['Image'] || '#', '_blank');
    }
  };

  // ------------------------------------
  // AUTH
  // ------------------------------------
  const [isAuthenticated, setIsAuthenticated] = useState(() => {
    try {
      return localStorage.getItem('app_auth') === 'true';
    } catch {
      return true;
    }
  });
  const [mobileNumber, setMobileNumber] = useState('');
  const [password, setPassword] = useState('');
  const [authError, setAuthError] = useState('');

  // ------------------------------------
  // DATA POOLS
  // ------------------------------------
  const [masterCatalog, setMasterCatalog] = useState([]);
  const [isLoadingCatalog, setIsLoadingCatalog] = useState(false);
  const [catalogLoadError, setCatalogLoadError] = useState('');

  const [myExcelData, setMyExcelData] = useState(() => {
    try {
      const saved = localStorage.getItem('app_curated_excel');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem('app_curated_excel', JSON.stringify(myExcelData));
    } catch (e) {}
  }, [myExcelData]);

  // ========================================================
  // 🚀 ROBUST AUTO-LOAD FOR DEPLOYED SERVERS
  // ========================================================
  const fetchAndLoadFile = async (url) => {
    const response = await fetch(url);
    if (!response.ok) return null;
    const arrayBuffer = await response.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(ws, { defval: '' });
    return rawData;
  };

  const autoLoadCatalog = async () => {
    // 1. Check local IndexedDB first
    let saved = await loadFromDB('master_catalog_pool');
    if (saved && Array.isArray(saved) && saved.length > 0) {
      setMasterCatalog(saved);
      return;
    }

    setIsLoadingCatalog(true);
    setCatalogLoadError('');

    // 2. Try loading common names from public folder
    const possibleFiles = ['/catalog.xlsx', '/Catalog.xlsx', '/data.xlsx', '/products.xlsx', '/catalog.csv'];
    let rawData = null;

    for (const filePath of possibleFiles) {
      try {
        rawData = await fetchAndLoadFile(filePath);
        if (rawData && rawData.length > 0) break;
      } catch (err) {
        // try next file
      }
    }

    if (rawData && rawData.length > 0) {
      const getVal = (row, possibleKeys) => {
        const rowKeys = Object.keys(row);
        for (const pKey of possibleKeys) {
          const matched = rowKeys.find(k => k.trim().toLowerCase() === pKey.toLowerCase());
          if (matched !== undefined && row[matched] !== undefined) {
            return String(row[matched]).trim();
          }
        }
        return '';
      };

      const normalized = rawData.map((row) => {
        const hindi = getVal(row, ['Hindi Name', 'HindiName', 'Hindi', 'regionalName']);
        const hinglish = getVal(row, ['Hinglish Name', 'HinglishName', 'Hinglish']);

        return {
          'Image': getVal(row, ['Image', 'imageUrl', 'image', 'img']),
          'Name': getVal(row, ['Name', 'name', 'productName', 'Title']),
          'Price': getVal(row, ['Price', 'lingPr', 'sellingPrice', 'price', 'sp']),
          'Original Price': getVal(row, ['Original Price', 'OriginalPrice', 'mrp', 'MRP']),
          'Quantity': getVal(row, ['Quantity', 'quantity', 'qty', 'weight', 'size']),
          'Sub-Category': getVal(row, ['Sub-Category', 'SubCategory', 'subcategory']),
          'Category': getVal(row, ['Category', 'category']),
          'Hindi Name': hindi,
          'Hinglish Name': hinglish,
          'Indian Category': getVal(row, ['Indian Category', 'IndianCategory']),
          'Indian Sub-Category': getVal(row, ['Indian Sub-Category', 'IndianSubCategory'])
        };
      });

      setMasterCatalog(normalized);
      saveToDB('master_catalog_pool', normalized);
      setCatalogLoadError('');
    } else {
      setCatalogLoadError('catalog.xlsx not found in public folder. Click "Import Dataset" to upload it once.');
    }

    setIsLoadingCatalog(false);
  };

  useEffect(() => {
    autoLoadCatalog();
  }, []);

  const updateCatalog = (newData) => {
    setMasterCatalog(newData);
    saveToDB('master_catalog_pool', newData);
    setCatalogLoadError('');
  };

  // Search & Pagination States
  const [searchQuery, setSearchQuery] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 7;
  const fileInputRef = useRef(null);

  // Modal State for Add / Edit
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editIndex, setEditIndex] = useState(null);
  const [formData, setFormData] = useState({
    'Image': '',
    'Name': '',
    'Price': '',
    'Original Price': '',
    'Quantity': '',
    'Sub-Category': '',
    'Category': '',
    'Hindi Name': '',
    'Hinglish Name': '',
    'Indian Category': '',
    'Indian Sub-Category': ''
  });

  const [fullscreenImage, setFullscreenImage] = useState(null);
  const [isUploadingToCloudinary, setIsUploadingToCloudinary] = useState(false);

  // Cloudinary Upload
  const handleCloudinaryUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!CLOUDINARY_CLOUD_NAME) {
      alert('Missing VITE_CLOUDINARY_CLOUD_NAME in .env file!');
      return;
    }

    setIsUploadingToCloudinary(true);
    try {
      const uploadData = new FormData();
      uploadData.append('file', file);

      if (CLOUDINARY_API_KEY && CLOUDINARY_API_SECRET) {
        const timestamp = Math.round(new Date().getTime() / 1000);
        const stringToSign = `timestamp=${timestamp}${CLOUDINARY_API_SECRET}`;
        const signature = await generateSHA1Signature(stringToSign);

        uploadData.append('api_key', CLOUDINARY_API_KEY);
        uploadData.append('timestamp', timestamp);
        uploadData.append('signature', signature);
      } else if (CLOUDINARY_UPLOAD_PRESET) {
        uploadData.append('upload_preset', CLOUDINARY_UPLOAD_PRESET);
        if (CLOUDINARY_API_KEY) uploadData.append('api_key', CLOUDINARY_API_KEY);
      } else {
        alert('Set VITE_CLOUDINARY_API_KEY & SECRET or VITE_CLOUDINARY_UPLOAD_PRESET in .env!');
        setIsUploadingToCloudinary(false);
        return;
      }

      const res = await fetch(`https://api.cloudinary.com/v1_1/${CLOUDINARY_CLOUD_NAME}/image/upload`, {
        method: 'POST',
        body: uploadData
      });

      const data = await res.json();
      if (data.secure_url) {
        setFormData(prev => ({ ...prev, 'Image': data.secure_url }));
        alert('Image uploaded successfully!');
      } else {
        alert('Upload Failed: ' + (data.error?.message || 'Check .env settings.'));
      }
    } catch (err) {
      console.error(err);
      alert('Error uploading to Cloudinary.');
    } finally {
      setIsUploadingToCloudinary(false);
      e.target.value = null;
    }
  };

  // Auth Methods
  const handleLogin = (e) => {
    e.preventDefault();
    if (!/^\d{10}$/.test(mobileNumber)) {
      setAuthError('Please enter a valid 10-digit mobile number.');
      return;
    }
    if (password.length < 4) {
      setAuthError('Password must be at least 4 characters.');
      return;
    }
    setAuthError('');
    setIsAuthenticated(true);
    try {
      localStorage.setItem('app_auth', 'true');
    } catch {}
  };

  const handleLogout = () => {
    setIsAuthenticated(false);
    try {
      localStorage.removeItem('app_auth');
    } catch {}
  };

  // Manual Excel Import (Permanent to IndexedDB)
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws, { defval: '' });

        if (rawData.length > 0) {
          const getVal = (row, possibleKeys) => {
            const rowKeys = Object.keys(row);
            for (const pKey of possibleKeys) {
              const matched = rowKeys.find(k => k.trim().toLowerCase() === pKey.toLowerCase());
              if (matched !== undefined && row[matched] !== undefined) {
                return String(row[matched]).trim();
              }
            }
            return '';
          };

          const normalized = rawData.map((row) => {
            const hindi = getVal(row, ['Hindi Name', 'HindiName', 'Hindi', 'regionalName']);
            const hinglish = getVal(row, ['Hinglish Name', 'HinglishName', 'Hinglish']);

            return {
              'Image': getVal(row, ['Image', 'imageUrl', 'image', 'img']),
              'Name': getVal(row, ['Name', 'name', 'productName', 'Title']),
              'Price': getVal(row, ['Price', 'lingPr', 'sellingPrice', 'price', 'sp']),
              'Original Price': getVal(row, ['Original Price', 'OriginalPrice', 'mrp', 'MRP']),
              'Quantity': getVal(row, ['Quantity', 'quantity', 'qty', 'weight', 'size']),
              'Sub-Category': getVal(row, ['Sub-Category', 'SubCategory', 'subcategory']),
              'Category': getVal(row, ['Category', 'category']),
              'Hindi Name': hindi,
              'Hinglish Name': hinglish,
              'Indian Category': getVal(row, ['Indian Category', 'IndianCategory']),
              'Indian Sub-Category': getVal(row, ['Indian Sub-Category', 'IndianSubCategory'])
            };
          });

          updateCatalog(normalized);
          alert(`Successfully loaded ${normalized.length} items permanently into this browser!`);
        }
      } catch (err) {
        console.error(err);
        alert('Could not read the file.');
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = null;
  };

  // Search across 33,000 items
  const searchResults = useMemo(() => {
    if (!searchQuery.trim()) return [];

    const queryTokens = searchQuery.trim().toLowerCase().split(/\s+/).filter(Boolean);
    const matches = [];

    for (let i = 0; i < masterCatalog.length; i++) {
      const row = masterCatalog[i];
      const hindi = (row['Hindi Name'] || '').toLowerCase();
      const hinglish = (row['Hinglish Name'] || '').toLowerCase();
      const name = (row['Name'] || '').toLowerCase();
      const subCat = (row['Sub-Category'] || '').toLowerCase();
      const indSubCat = (row['Indian Sub-Category'] || '').toLowerCase();

      const combined = `${name} ${hindi} ${hinglish} ${subCat} ${indSubCat}`;

      if (queryTokens.every((token) => combined.includes(token))) {
        matches.push(row);
        if (matches.length >= 24) break;
      }
    }
    return matches;
  }, [masterCatalog, searchQuery]);

  const isItemAdded = (itemName) => {
    return myExcelData.some(item => item['Name'].trim().toLowerCase() === itemName.trim().toLowerCase());
  };

  const handleAddToMyExcel = (product) => {
    setMyExcelData(prev => [product, ...prev]);
  };

  // Export
  const handleDownloadExcel = () => {
    if (myExcelData.length === 0) {
      alert('Your Excel list is empty! Search and add products first.');
      return;
    }

    const ws = XLSX.utils.json_to_sheet(myExcelData, { header: COLUMNS });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'My_Products');
    XLSX.writeFile(wb, `My_Excel_Database_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // CRUD
  const openEditModal = (row, index) => {
    setEditIndex(index);
    setFormData({ ...row });
    setIsModalOpen(true);
  };

  const openAddModal = () => {
    setEditIndex(null);
    const emptyForm = {};
    COLUMNS.forEach(c => emptyForm[c] = '');
    setFormData(emptyForm);
    setIsModalOpen(true);
  };

  const handleSaveModal = async (e) => {
    e.preventDefault();
    let finalHindi = formData['Hindi Name'];
    if (!finalHindi && formData['Hinglish Name']) {
      finalHindi = await convertHinglishToHindi(formData['Hinglish Name']);
    }

    const finalData = { ...formData, 'Hindi Name': finalHindi || '' };

    if (editIndex !== null) {
      const updated = [...myExcelData];
      updated[editIndex] = finalData;
      setMyExcelData(updated);
    } else {
      setMyExcelData([finalData, ...myExcelData]);
    }
    setIsModalOpen(false);
  };

  const handleDeleteRow = (index) => {
    if (window.confirm('Delete this product from your Excel list?')) {
      setMyExcelData(myExcelData.filter((_, idx) => idx !== index));
    }
  };

  const totalPages = Math.ceil(myExcelData.length / rowsPerPage) || 1;
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return myExcelData.slice(start, start + rowsPerPage);
  }, [myExcelData, currentPage]);

  // ========================================================
  // 🌟 INDIVIDUAL PRODUCT VIEW (NEW TAB)
  // ========================================================
  if (individualProduct) {
    const added = isItemAdded(individualProduct['Name']);

    return (
      <div className="min-h-screen bg-slate-50 text-slate-800 font-sans p-4 sm:p-8 flex flex-col items-center">
        <div className="max-w-4xl w-full bg-white border border-slate-200 rounded-3xl p-6 sm:p-10 shadow-xl space-y-6">
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <button
              onClick={() => {
                if (window.opener) {
                  window.close();
                } else {
                  window.location.href = window.location.pathname;
                }
              }}
              className="flex items-center gap-2 text-xs font-semibold text-slate-600 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 px-3.5 py-2 rounded-xl transition cursor-pointer"
            >
              <ArrowLeft className="w-4 h-4" />
              <span>Back / Close Tab</span>
            </button>

            <span className="text-xs bg-emerald-50 text-emerald-700 border border-emerald-200 font-semibold px-3 py-1 rounded-full">
              Individual Product View
            </span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-8 items-start">
            <div 
              className="relative bg-slate-50 border border-slate-200 rounded-3xl p-6 flex items-center justify-center cursor-pointer group overflow-hidden"
              onClick={() => individualProduct['Image'] && setFullscreenImage(individualProduct['Image'])}
              title="Click for Fullscreen Zoom"
            >
              {individualProduct['Image'] ? (
                <img
                  src={individualProduct['Image']}
                  alt=""
                  className="w-full max-h-80 object-contain rounded-2xl group-hover:scale-105 transition duration-300"
                  onError={(e) => { e.target.style.display = 'none'; }}
                />
              ) : (
                <div className="w-48 h-48 flex items-center justify-center text-slate-300">
                  <ImageIcon className="w-16 h-16" />
                </div>
              )}
              <div className="absolute bottom-3 right-3 bg-black/60 text-white text-[11px] px-2.5 py-1 rounded-lg flex items-center gap-1.5 opacity-0 group-hover:opacity-100 transition">
                <Maximize2 className="w-3.5 h-3.5" />
                <span>Tap to Zoom</span>
              </div>
            </div>

            <div className="space-y-4">
              <div>
                <h1 className="text-xl sm:text-2xl font-bold text-slate-900 leading-tight">
                  {individualProduct['Name']}
                </h1>
                <p className="text-sm font-semibold text-emerald-600 mt-1">
                  {individualProduct['Hindi Name'] || '-'}
                </p>
                <p className="text-xs font-medium text-blue-600 mt-0.5">
                  {individualProduct['Hinglish Name'] || '-'}
                </p>
              </div>

              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 flex items-center justify-between">
                <div>
                  <span className="text-xs text-slate-400 block font-semibold uppercase">Selling Price</span>
                  <div className="flex items-baseline gap-2">
                    <span className="text-2xl font-bold text-emerald-600">₹{individualProduct['Price']}</span>
                    {individualProduct['Original Price'] && (
                      <span className="text-sm text-slate-400 line-through">₹{individualProduct['Original Price']}</span>
                    )}
                  </div>
                </div>

                <div className="text-right">
                  <span className="text-xs text-slate-400 block font-semibold uppercase">Quantity</span>
                  <span className="text-sm font-bold text-slate-800">{individualProduct['Quantity'] || 'N/A'}</span>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3 text-xs">
                <div className="bg-slate-50 p-3 rounded-xl border border-slate-200">
                  <span className="text-[10px] text-slate-400 font-semibold uppercase block">Category</span>
                  <span className="font-semibold text-slate-800">{individualProduct['Category'] || '-'}</span>
                </div>
                <div className="bg-slate-50 p-3 rounded-xl border border-slate-200">
                  <span className="text-[10px] text-slate-400 font-semibold uppercase block">Sub-Category</span>
                  <span className="font-semibold text-slate-800">{individualProduct['Sub-Category'] || '-'}</span>
                </div>
                <div className="bg-slate-50 p-3 rounded-xl border border-slate-200">
                  <span className="text-[10px] text-slate-400 font-semibold uppercase block">Indian Category</span>
                  <span className="font-semibold text-slate-800">{individualProduct['Indian Category'] || '-'}</span>
                </div>
                <div className="bg-slate-50 p-3 rounded-xl border border-slate-200">
                  <span className="text-[10px] text-slate-400 font-semibold uppercase block">Indian Sub-Category</span>
                  <span className="font-semibold text-slate-800">{individualProduct['Indian Sub-Category'] || '-'}</span>
                </div>
              </div>

              <div className="pt-2">
                {added ? (
                  <div className="w-full py-3 bg-emerald-50 border border-emerald-200 text-emerald-700 font-bold rounded-2xl flex items-center justify-center gap-2 shadow-xs text-sm">
                    <Check className="w-5 h-5" />
                    <span>Already in My Excel Database</span>
                  </div>
                ) : (
                  <button
                    onClick={() => handleAddToMyExcel(individualProduct)}
                    className="w-full py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-2xl flex items-center justify-center gap-2 shadow-lg shadow-emerald-600/20 text-sm transition active:scale-98 cursor-pointer"
                  >
                    <Plus className="w-5 h-5" />
                    <span>+ Add to My Excel Database</span>
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>

        {fullscreenImage && (
          <div 
            className="fixed inset-0 z-60 bg-black/90 backdrop-blur-md flex items-center justify-center p-4 cursor-pointer"
            onClick={() => setFullscreenImage(null)}
          >
            <button
              onClick={() => setFullscreenImage(null)}
              className="absolute top-4 right-4 p-2 bg-slate-800 text-white rounded-full hover:bg-slate-700 transition"
            >
              <X className="w-6 h-6" />
            </button>
            <img
              src={fullscreenImage}
              alt=""
              className="max-w-full max-h-[90vh] object-contain rounded-2xl shadow-2xl"
              onClick={(e) => e.stopPropagation()}
            />
          </div>
        )}
      </div>
    );
  }

  // ========================================================
  // ⚪ LOGIN SCREEN
  // ========================================================
  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4">
        <div className="max-w-md w-full bg-white border border-slate-200 rounded-3xl p-6 sm:p-8 shadow-xl">
          <div className="text-center mb-6">
            <div className="inline-flex p-3.5 bg-emerald-50 text-emerald-600 rounded-2xl mb-3 border border-emerald-100">
              <FileSpreadsheet className="w-8 h-8" />
            </div>
            <h2 className="text-xl font-bold text-slate-800 tracking-tight">Excel Database Creator</h2>
            <p className="text-slate-500 text-xs mt-1">Light & Crisp UI • Auto-Loaded Catalog</p>
          </div>

          {authError && (
            <div className="mb-4 flex items-center gap-2 p-3 bg-red-50 border border-red-200 rounded-xl text-red-600 text-xs">
              <AlertCircle className="w-4 h-4 shrink-0" />
              <span>{authError}</span>
            </div>
          )}

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase tracking-wider mb-1.5">
                Mobile Number
              </label>
              <div className="relative">
                <Phone className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="tel"
                  placeholder="10-digit mobile number"
                  maxLength={10}
                  value={mobileNumber}
                  onChange={(e) => setMobileNumber(e.target.value.replace(/\D/g, ''))}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl pl-10 pr-4 py-2.5 text-sm text-slate-800 placeholder-slate-400 focus:outline-none focus:border-emerald-500 focus:bg-white transition"
                  required
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase tracking-wider mb-1.5">
                Password
              </label>
              <div className="relative">
                <Lock className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="password"
                  placeholder="••••••••"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl pl-11 pr-4 py-2.5 text-sm text-slate-800 placeholder-slate-400 focus:outline-none focus:border-emerald-500 focus:bg-white transition"
                  required
                />
              </div>
            </div>

            <button
              type="submit"
              className="w-full mt-2 py-2.5 px-4 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-semibold rounded-xl transition duration-150 cursor-pointer shadow-lg shadow-emerald-600/20"
            >
              Sign In
            </button>
          </form>
        </div>
      </div>
    );
  }

  // ========================================================
  // ⚪ MAIN WORKSPACE
  // ========================================================
  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col font-sans pb-24 md:pb-8">
      {/* Top Navbar */}
      <header className="border-b border-slate-200 bg-white/90 backdrop-blur-md sticky top-0 z-30 shadow-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 h-14 sm:h-16 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="p-1.5 sm:p-2 bg-emerald-50 text-emerald-600 rounded-xl border border-emerald-100">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <h1 className="font-bold text-sm sm:text-base text-slate-800 leading-none">Excel Database Studio</h1>
              <p className="text-[10px] sm:text-[11px] text-slate-500 mt-0.5 flex items-center gap-1">
                {isLoadingCatalog ? (
                  <span className="text-blue-600 font-medium animate-pulse flex items-center gap-1">
                    <Loader2 className="w-3 h-3 animate-spin" />
                    Auto-loading 33,000+ catalog...
                  </span>
                ) : (
                  <>
                    Catalog: <strong className="text-emerald-600">{masterCatalog.length.toLocaleString()} items</strong>
                    {masterCatalog.length === 0 && (
                      <button
                        onClick={autoLoadCatalog}
                        className="ml-2 text-blue-600 underline flex items-center gap-1 cursor-pointer"
                      >
                        <RefreshCw className="w-2.5 h-2.5" />
                        <span>Retry Load</span>
                      </button>
                    )}
                  </>
                )}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <span className="hidden sm:inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] bg-slate-100 border border-slate-200 text-slate-600 rounded-xl">
              <UploadCloud className="w-3 h-3 text-emerald-600" />
              <span>Cloud: <strong className="text-slate-800">{CLOUDINARY_CLOUD_NAME}</strong></span>
            </span>

            <button
              onClick={handleLogout}
              className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs text-slate-600 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 rounded-xl border border-slate-200 transition cursor-pointer"
            >
              <LogOut className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Logout</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-3 sm:p-6 space-y-4 sm:space-y-6">
        
        {/* Notice banner if 0 items are loaded */}
        {masterCatalog.length === 0 && !isLoadingCatalog && (
          <div className="p-3 bg-amber-50 border border-amber-200 rounded-2xl text-amber-800 text-xs flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 shadow-xs">
            <div className="flex items-center gap-2">
              <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
              <span>
                <strong>Suggestion catalog is currently 0:</strong> Click "Import Dataset" to upload your 33,000 items once, or verify <code>public/catalog.xlsx</code> on GitHub.
              </span>
            </div>
            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3 py-1 bg-amber-600 hover:bg-amber-700 text-white font-semibold rounded-lg shrink-0 cursor-pointer"
            >
              Upload Dataset Now
            </button>
          </div>
        )}

        {/* Search & Actions Bar */}
        <div className="flex flex-col lg:flex-row gap-3 justify-between items-stretch lg:items-center bg-white border border-slate-200 rounded-2xl p-3 sm:p-4 shadow-sm">
          <div className="relative flex-1">
            <Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-emerald-600" />
            <input
              type="text"
              placeholder={`Search in Hindi (गेहूं, चाय) or Hinglish (red label, atta)...`}
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full bg-slate-50 border border-slate-200 rounded-xl pl-10 pr-9 py-2 sm:py-2.5 text-xs sm:text-sm text-slate-800 placeholder-slate-400 focus:outline-none focus:border-emerald-500 focus:bg-white focus:ring-1 focus:ring-emerald-500 transition shadow-xs"
            />
            {searchQuery && (
              <button
                onClick={() => setSearchQuery('')}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <input
              type="file"
              accept=".xlsx, .xls, .csv"
              ref={fileInputRef}
              onChange={handleFileUpload}
              className="hidden"
            />

            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-3 py-2 text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl border border-slate-200 transition cursor-pointer font-medium"
            >
              <Upload className="w-3.5 h-3.5 text-emerald-600" />
              <span>Import Dataset</span>
            </button>

            <button
              onClick={openAddModal}
              className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-3 py-2 text-xs bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl shadow-md shadow-emerald-600/20 font-medium transition cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Add Custom</span>
            </button>

            <button
              onClick={handleDownloadExcel}
              className="hidden sm:flex items-center gap-1.5 px-3 py-2 text-xs bg-blue-600 hover:bg-blue-700 text-white rounded-xl shadow-md shadow-blue-600/20 font-medium transition cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Download Excel ({myExcelData.length})</span>
            </button>
          </div>
        </div>

        {/* Search Results */}
        {searchQuery.trim().length > 0 && (
          <div className="bg-white border border-emerald-200 rounded-2xl p-3 sm:p-4 space-y-3 shadow-md">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <div className="flex items-center gap-1.5">
                <Sparkles className="w-3.5 h-3.5 text-emerald-600" />
                <h3 className="font-bold text-xs sm:text-sm text-slate-800">
                  Found {searchResults.length} matches:
                </h3>
              </div>
              <span className="text-[11px] text-slate-500 hidden sm:inline">
                Click 👁️ to open in <strong>New Tab</strong> or <strong>"+ Add"</strong> to select
              </span>
            </div>

            {searchResults.length > 0 ? (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-2.5 sm:gap-3 max-h-96 overflow-y-auto pr-1">
                {searchResults.map((item, idx) => {
                  const added = isItemAdded(item['Name']);

                  return (
                    <div
                      key={idx}
                      className="bg-slate-50/70 border border-slate-200 rounded-xl p-3 flex flex-col justify-between hover:border-slate-300 hover:bg-white transition gap-2 shadow-xs"
                    >
                      <div className="flex items-start gap-2.5">
                        <div 
                          className="relative cursor-pointer group shrink-0" 
                          onClick={() => handleOpenIndividualTab(item)}
                          title="Open in New Tab"
                        >
                          {item['Image'] ? (
                            <img
                              src={item['Image']}
                              alt=""
                              className="w-12 h-12 object-cover rounded-lg border border-slate-200 bg-white group-hover:opacity-85 transition"
                              onError={(e) => { e.target.style.display = 'none'; }}
                            />
                          ) : (
                            <div className="w-12 h-12 rounded-lg bg-slate-200 border border-slate-200 flex items-center justify-center text-slate-400">
                              <ImageIcon className="w-5 h-5" />
                            </div>
                          )}
                          <ExternalLink className="w-3 h-3 text-white absolute bottom-1 right-1 opacity-0 group-hover:opacity-100 bg-black/60 rounded p-0.5 transition" />
                        </div>

                        <div className="min-w-0 flex-1">
                          <p 
                            className="text-xs font-semibold text-slate-800 line-clamp-1 hover:text-emerald-600 cursor-pointer" 
                            onClick={() => handleOpenIndividualTab(item)}
                            title="Open in New Tab"
                          >
                            {item['Name']}
                          </p>
                          <p className="text-[11px] text-emerald-600 font-medium line-clamp-1">
                            {item['Hindi Name'] || '-'}
                          </p>
                          <p className="text-[10px] text-blue-600 line-clamp-1">
                            {item['Hinglish Name'] || '-'}
                          </p>
                        </div>
                      </div>

                      <div className="flex items-center justify-between pt-2 border-t border-slate-200/80 text-xs">
                        <div className="flex items-baseline gap-1.5">
                          <span className="font-bold text-emerald-600">₹{item['Price']}</span>
                          {item['Original Price'] && (
                            <span className="text-[10px] text-slate-400 line-through">₹{item['Original Price']}</span>
                          )}
                          <span className="text-[10px] text-slate-500">{item['Quantity']}</span>
                        </div>

                        <div className="flex items-center gap-1.5">
                          <button
                            onClick={() => handleOpenIndividualTab(item)}
                            className="p-1.5 bg-white border border-slate-200 hover:bg-slate-100 text-slate-600 rounded-lg transition cursor-pointer flex items-center gap-1"
                            title="Open product details in a new tab"
                          >
                            <Eye className="w-3.5 h-3.5" />
                          </button>

                          {added ? (
                            <div className="flex items-center gap-1 text-[11px] text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-1 rounded-lg">
                              <Check className="w-3 h-3" />
                              <span>Added</span>
                            </div>
                          ) : (
                            <button
                              onClick={() => handleAddToMyExcel(item)}
                              className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-medium flex items-center gap-1 transition cursor-pointer shadow-xs active:scale-95"
                            >
                              <Plus className="w-3 h-3" />
                              <span>Add</span>
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="py-6 text-center text-slate-400 text-xs">
                {masterCatalog.length === 0 
                  ? 'Catalog is empty. Please upload dataset or place catalog.xlsx in public folder.' 
                  : `No products match "${searchQuery}".`}
              </div>
            )}
          </div>
        )}

        {/* Database Table */}
        <div className="space-y-3">
          <div className="flex items-center justify-between">
            <div>
              <h2 className="text-base sm:text-lg font-bold text-slate-800 tracking-tight">
                My Final Excel Database ({myExcelData.length})
              </h2>
              <p className="text-[11px] text-slate-500">
                Only these {myExcelData.length} items will be downloaded.
              </p>
            </div>

            {myExcelData.length > 0 && (
              <button
                onClick={() => {
                  if (window.confirm('Clear all products from your Excel database?')) setMyExcelData([]);
                }}
                className="text-xs text-red-600 hover:underline cursor-pointer"
              >
                Clear All
              </button>
            )}
          </div>

          {/* Mobile View */}
          <div className="block md:hidden space-y-2">
            {paginatedData.length > 0 ? (
              paginatedData.map((row, rIdx) => {
                const originalIndex = (currentPage - 1) * rowsPerPage + rIdx;
                return (
                  <div key={rIdx} className="bg-white border border-slate-200 rounded-xl p-3 space-y-2 shadow-xs">
                    <div className="flex items-start gap-3">
                      <div 
                        className="cursor-pointer relative group shrink-0"
                        onClick={() => handleOpenIndividualTab(row)}
                      >
                        {row['Image'] ? (
                          <img
                            src={row['Image']}
                            alt=""
                            className="w-12 h-12 object-cover rounded-lg border border-slate-200 bg-slate-50"
                            onError={(e) => { e.target.style.display = 'none'; }}
                          />
                        ) : (
                          <div className="w-12 h-12 rounded-lg bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-400">
                            <ImageIcon className="w-5 h-5" />
                          </div>
                        )}
                        <ExternalLink className="w-3 h-3 text-white absolute bottom-1 right-1 opacity-0 group-hover:opacity-100 bg-black/60 rounded p-0.5" />
                      </div>

                      <div className="flex-1 min-w-0">
                        <h4 
                          className="text-xs font-semibold text-slate-800 line-clamp-1 hover:text-emerald-600 cursor-pointer"
                          onClick={() => handleOpenIndividualTab(row)}
                        >
                          {row['Name']}
                        </h4>
                        <div className="flex items-center gap-1.5 mt-0.5">
                          <span className="text-[11px] text-emerald-600 font-medium line-clamp-1">{row['Hindi Name'] || '-'}</span>
                          <span className="text-slate-300">•</span>
                          <span className="text-[10px] text-blue-600 line-clamp-1">{row['Hinglish Name'] || '-'}</span>
                        </div>
                        <div className="flex items-center gap-2 mt-1 text-xs">
                          <span className="font-bold text-emerald-600">₹{row['Price']}</span>
                          {row['Original Price'] && (
                            <span className="text-[10px] text-slate-400 line-through">₹{row['Original Price']}</span>
                          )}
                          <span className="text-[10px] text-slate-500">{row['Quantity']}</span>
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center justify-between pt-2 border-t border-slate-100 text-xs">
                      <span className="text-[10px] text-slate-400">#{originalIndex + 1} • {row['Sub-Category'] || 'Item'}</span>
                      <div className="flex items-center gap-1.5">
                        <button
                          onClick={() => handleOpenIndividualTab(row)}
                          className="p-1.5 bg-slate-100 text-slate-600 rounded-lg hover:bg-slate-200"
                          title="Open in New Tab"
                        >
                          <Eye className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => openEditModal(row, originalIndex)}
                          className="p-1.5 bg-slate-100 text-slate-600 rounded-lg hover:bg-slate-200"
                        >
                          <Edit3 className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => handleDeleteRow(originalIndex)}
                          className="p-1.5 bg-red-50 text-red-600 rounded-lg hover:bg-red-100"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })
            ) : (
              <div className="py-10 text-center text-slate-400 text-xs bg-white rounded-2xl border border-slate-200">
                Your database is empty. Search above to add items!
              </div>
            )}
          </div>

          {/* Desktop Table View */}
          <div className="hidden md:block bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-xs">
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-slate-700">
                <thead className="bg-slate-100 border-b border-slate-200 text-slate-600 uppercase tracking-wider text-[11px]">
                  <tr>
                    <th className="py-3 px-3 w-10 text-center">#</th>
                    <th className="py-3 px-3 font-semibold whitespace-nowrap">Image</th>
                    <th className="py-3 px-3 font-semibold whitespace-nowrap min-w-[180px]">Name</th>
                    <th className="py-3 px-3 font-semibold whitespace-nowrap">Price</th>
                    <th className="py-3 px-3 font-semibold whitespace-nowrap">Original Price</th>
                    <th className="py-3 px-3 font-semibold whitespace-nowrap">Quantity</th>
                    <th className="py-3 px-3 font-semibold whitespace-nowrap">Sub-Category</th>
                    <th className="py-3 px-3 font-semibold whitespace-nowrap min-w-[140px]">Category</th>
                    <th className="py-3 px-3 font-semibold whitespace-nowrap text-emerald-700 min-w-[130px]">Hindi Name</th>
                    <th className="py-3 px-3 font-semibold whitespace-nowrap text-blue-700 min-w-[150px]">Hinglish Name</th>
                    <th className="py-3 px-3 font-semibold whitespace-nowrap min-w-[120px]">Indian Category</th>
                    <th className="py-3 px-3 font-semibold whitespace-nowrap min-w-[120px]">Indian Sub-Category</th>
                    <th className="py-3 px-3 text-right pr-4 whitespace-nowrap">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {paginatedData.length > 0 ? (
                    paginatedData.map((row, rIdx) => {
                      const originalIndex = (currentPage - 1) * rowsPerPage + rIdx;
                      return (
                        <tr key={rIdx} className="hover:bg-slate-50 transition">
                          <td className="py-3 px-3 text-center text-slate-400">
                            {originalIndex + 1}
                          </td>

                          <td className="py-3 px-3 whitespace-nowrap">
                            <div 
                              className="cursor-pointer relative group inline-block"
                              onClick={() => handleOpenIndividualTab(row)}
                              title="Open in New Tab"
                            >
                              {row['Image'] ? (
                                <img
                                  src={row['Image']}
                                  alt=""
                                  className="w-9 h-9 object-cover rounded-lg border border-slate-200 bg-slate-50 group-hover:opacity-85 transition"
                                  onError={(e) => { e.target.style.display = 'none'; }}
                                />
                              ) : (
                                <div className="w-9 h-9 rounded-lg bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-400">
                                  <ImageIcon className="w-4 h-4" />
                                </div>
                              )}
                              <ExternalLink className="w-2.5 h-2.5 text-white absolute bottom-1 right-1 opacity-0 group-hover:opacity-100 bg-black/60 rounded p-0.5" />
                            </div>
                          </td>

                          <td 
                            className="py-3 px-3 font-medium text-slate-800 whitespace-nowrap max-w-xs truncate hover:text-emerald-600 cursor-pointer" 
                            onClick={() => handleOpenIndividualTab(row)}
                            title="Open in New Tab"
                          >
                            {row['Name']}
                          </td>
                          <td className="py-3 px-3 font-bold text-emerald-600 whitespace-nowrap">₹{row['Price']}</td>
                          <td className="py-3 px-3 text-slate-400 whitespace-nowrap">₹{row['Original Price']}</td>
                          <td className="py-3 px-3 whitespace-nowrap text-slate-600">{row['Quantity'] || '-'}</td>
                          <td className="py-3 px-3 whitespace-nowrap text-slate-600">{row['Sub-Category'] || '-'}</td>
                          <td className="py-3 px-3 whitespace-nowrap text-slate-600">{row['Category'] || '-'}</td>
                          <td className="py-3 px-3 font-semibold text-emerald-700 whitespace-nowrap">{row['Hindi Name'] || '-'}</td>
                          <td className="py-3 px-3 font-semibold text-blue-700 whitespace-nowrap">{row['Hinglish Name'] || '-'}</td>
                          <td className="py-3 px-3 whitespace-nowrap text-slate-600">{row['Indian Category'] || '-'}</td>
                          <td className="py-3 px-3 whitespace-nowrap text-slate-600">{row['Indian Sub-Category'] || '-'}</td>

                          <td className="py-3 px-3 text-right pr-4 whitespace-nowrap">
                            <div className="flex items-center justify-end gap-1.5">
                              <button
                                onClick={() => handleOpenIndividualTab(row)}
                                className="p-1.5 hover:bg-slate-100 text-slate-500 hover:text-emerald-600 rounded-lg transition cursor-pointer"
                                title="Open in Individual New Tab"
                              >
                                <Eye className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => openEditModal(row, originalIndex)}
                                className="p-1.5 hover:bg-slate-100 text-slate-500 hover:text-slate-800 rounded-lg transition cursor-pointer"
                                title="Edit"
                              >
                                <Edit3 className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => handleDeleteRow(originalIndex)}
                                className="p-1.5 hover:bg-red-50 text-slate-500 hover:text-red-600 rounded-lg transition cursor-pointer"
                                title="Delete"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={13} className="py-14 text-center text-slate-400">
                        Your Excel database is empty. Search above and click "+ Add" to select products!
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* Pagination */}
          {myExcelData.length > 0 && (
            <div className="p-3 bg-white border border-slate-200 rounded-xl flex items-center justify-between text-xs text-slate-600 shadow-xs">
              <div>
                Showing {(currentPage - 1) * rowsPerPage + 1} to{' '}
                {Math.min(currentPage * rowsPerPage, myExcelData.length)} of {myExcelData.length}
              </div>

              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                  className="p-1.5 rounded-lg border border-slate-200 disabled:opacity-30 hover:bg-slate-100 transition cursor-pointer"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <span className="font-medium text-slate-800">
                  {currentPage} / {totalPages}
                </span>
                <button
                  onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                  disabled={currentPage === totalPages}
                  className="p-1.5 rounded-lg border border-slate-200 disabled:opacity-30 hover:bg-slate-100 transition cursor-pointer"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}
        </div>
      </main>

      {/* Floating Bottom Bar on Mobile */}
      <div className="sm:hidden fixed bottom-3 left-3 right-3 bg-white/95 border border-slate-200 backdrop-blur-md rounded-2xl p-3 shadow-xl flex items-center justify-between z-40">
        <div>
          <p className="text-[10px] text-slate-500">My Excel List:</p>
          <p className="text-xs font-bold text-emerald-600">{myExcelData.length} products added</p>
        </div>
        <button
          onClick={handleDownloadExcel}
          className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5 shadow-md shadow-emerald-600/20 cursor-pointer active:scale-95"
        >
          <Download className="w-3.5 h-3.5" />
          <span>Download Excel</span>
        </button>
      </div>

      {/* Lightbox Zoom */}
      {fullscreenImage && (
        <div 
          className="fixed inset-0 z-60 bg-black/90 backdrop-blur-md flex items-center justify-center p-4 cursor-pointer"
          onClick={() => setFullscreenImage(null)}
        >
          <button
            onClick={() => setFullscreenImage(null)}
            className="absolute top-4 right-4 p-2 bg-slate-800/80 text-white rounded-full hover:bg-slate-700 transition cursor-pointer"
          >
            <X className="w-6 h-6" />
          </button>
          <img
            src={fullscreenImage}
            alt=""
            className="max-w-full max-h-[90vh] object-contain rounded-2xl shadow-2xl cursor-default"
            onClick={(e) => e.stopPropagation()}
          />
        </div>
      )}

      {/* Modal: Add/Edit Product */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-3 sm:p-4">
          <div className="bg-white border border-slate-200 rounded-3xl w-full max-w-xl overflow-hidden shadow-2xl">
            <div className="p-4 sm:p-5 border-b border-slate-100 flex items-center justify-between">
              <h3 className="font-bold text-slate-800 text-sm sm:text-base">
                {editIndex !== null ? 'Edit Product' : 'Add Custom Product'}
              </h3>
              <button onClick={() => setIsModalOpen(false)} className="text-slate-400 hover:text-slate-600 cursor-pointer p-1">
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleSaveModal} className="p-4 sm:p-6 space-y-3.5 max-h-[75vh] overflow-y-auto text-xs">
              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="font-semibold text-slate-600 uppercase">Image URL / Device Upload</label>
                  <span className="text-[10px] text-emerald-600">
                    Cloudinary active ({CLOUDINARY_CLOUD_NAME})
                  </span>
                </div>

                <div className="flex gap-2">
                  <input
                    type="text"
                    placeholder="https://res.cloudinary.com/..."
                    value={formData['Image']}
                    onChange={(e) => setFormData({ ...formData, 'Image': e.target.value })}
                    className="flex-1 bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:border-emerald-500 focus:bg-white"
                  />

                  <label className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-semibold cursor-pointer flex items-center gap-1.5 shrink-0 transition active:scale-95 shadow-xs">
                    {isUploadingToCloudinary ? (
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                    ) : (
                      <UploadCloud className="w-3.5 h-3.5" />
                    )}
                    <span>{isUploadingToCloudinary ? 'Uploading...' : 'Upload Image'}</span>
                    <input
                      type="file"
                      accept="image/*"
                      onChange={handleCloudinaryUpload}
                      disabled={isUploadingToCloudinary}
                      className="hidden"
                    />
                  </label>
                </div>

                {formData['Image'] && (
                  <div className="mt-2 flex items-center gap-2 bg-slate-50 p-2 rounded-xl border border-slate-200">
                    <img 
                      src={formData['Image']} 
                      alt="" 
                      className="w-10 h-10 object-cover rounded-lg border border-slate-200 cursor-pointer"
                      onClick={() => setFullscreenImage(formData['Image'])}
                    />
                    <div className="min-w-0 flex-1">
                      <span className="text-emerald-700 text-[11px] font-semibold block">✓ Image linked</span>
                      <span className="text-slate-400 text-[10px] truncate block">{formData['Image']}</span>
                    </div>
                  </div>
                )}
              </div>

              <div>
                <label className="block font-semibold text-slate-600 uppercase mb-1">Name</label>
                <input
                  type="text"
                  required
                  placeholder="Product Name"
                  value={formData['Name']}
                  onChange={(e) => setFormData({ ...formData, 'Name': e.target.value })}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:border-emerald-500 focus:bg-white"
                />
              </div>

              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="block font-semibold text-slate-600 uppercase mb-1">Price</label>
                  <input
                    type="number"
                    step="any"
                    required
                    placeholder="130"
                    value={formData['Price']}
                    onChange={(e) => setFormData({ ...formData, 'Price': e.target.value })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:border-emerald-500 focus:bg-white"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-600 uppercase mb-1">Original Price</label>
                  <input
                    type="number"
                    step="any"
                    placeholder="150"
                    value={formData['Original Price']}
                    onChange={(e) => setFormData({ ...formData, 'Original Price': e.target.value })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:border-emerald-500 focus:bg-white"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-600 uppercase mb-1">Quantity</label>
                  <input
                    type="text"
                    placeholder="500 g"
                    value={formData['Quantity']}
                    onChange={(e) => setFormData({ ...formData, 'Quantity': e.target.value })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:border-emerald-500 focus:bg-white"
                  />
                </div>
              </div>

              {/* Hinglish to Hindi Auto Converter */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 p-3 bg-slate-50 rounded-2xl border border-slate-200">
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="font-semibold text-blue-700 uppercase">Hinglish Name</label>
                    <span className="text-[10px] text-slate-500">Type in English</span>
                  </div>
                  <input
                    type="text"
                    placeholder="e.g. kapoor container, gehu ka atta"
                    value={formData['Hinglish Name']}
                    onChange={async (e) => {
                      const val = e.target.value;
                      setFormData(prev => ({ ...prev, 'Hinglish Name': val }));

                      const hindiResult = await convertHinglishToHindi(val);
                      if (hindiResult) {
                        setFormData(prev => ({ ...prev, 'Hindi Name': hindiResult }));
                      }
                    }}
                    className="w-full bg-white border border-slate-200 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:border-blue-500"
                  />
                </div>

                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="font-semibold text-emerald-700 uppercase">Hindi Name</label>
                    <span className="text-[10px] text-emerald-600 font-medium">✓ Auto-converted</span>
                  </div>
                  <input
                    type="text"
                    placeholder="कपूर कंटेनर (Auto-generated)"
                    value={formData['Hindi Name']}
                    onChange={(e) => setFormData(prev => ({ ...prev, 'Hindi Name': e.target.value }))}
                    className="w-full bg-white border border-slate-200 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:border-emerald-500"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-600 uppercase mb-1">Sub-Category</label>
                  <input
                    type="text"
                    placeholder="Pooja & Worship Needs"
                    value={formData['Sub-Category']}
                    onChange={(e) => setFormData({ ...formData, 'Sub-Category': e.target.value })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:border-emerald-500 focus:bg-white"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-600 uppercase mb-1">Category</label>
                  <input
                    type="text"
                    placeholder="Home Needs"
                    value={formData['Category']}
                    onChange={(e) => setFormData({ ...formData, 'Category': e.target.value })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:border-emerald-500 focus:bg-white"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-neutral-400 uppercase mb-1">Indian Category</label>
                  <input
                    type="text"
                    placeholder="Other"
                    value={formData['Indian Category']}
                    onChange={(e) => setFormData({ ...formData, 'Indian Category': e.target.value })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:border-emerald-500 focus:bg-white"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-neutral-400 uppercase mb-1">Indian Sub-Category</label>
                  <input
                    type="text"
                    placeholder="अन्य"
                    value={formData['Indian Sub-Category']}
                    onChange={(e) => setFormData({ ...formData, 'Indian Sub-Category': e.target.value })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:border-emerald-500 focus:bg-white"
                  />
                </div>
              </div>

              <div className="pt-3 flex gap-2 justify-end border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setIsModalOpen(false)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl transition cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-semibold rounded-xl transition cursor-pointer shadow-md shadow-emerald-600/20"
                >
                  {editIndex !== null ? 'Update Record' : 'Add to My Excel'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}