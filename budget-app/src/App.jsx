import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  PieChart, Pie, Cell, BarChart, Bar, XAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer, LabelList 
} from 'recharts';
import { 
  Plus, Wallet, Home, PieChart as ChartIcon, Settings, Trash2, ArrowDownToLine, ArrowUpFromLine, ChevronRight, X, Bot, Target, Download, Edit2, Mic
} from 'lucide-react';
import { v4 as uuidv4 } from 'uuid';
import { format, parseISO, isSameMonth, startOfMonth, endOfMonth, eachDayOfInterval, isSunday } from 'date-fns';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

const DEFAULT_CATEGORIES = [
  { id: 'savings', name: 'Запас ва Сармоя', color: '#10b981', tags: [] },
  { id: 'motivation', name: 'Мотивация (Захира)', color: '#059669', tags: ['пепси ичмадим', 'бекорга олмадим', 'захира', 'захирага', 'zaxira'] },
  { id: 'badhabits', name: 'Зарарли одатлар', color: '#dc2626', tags: ['пепси', 'энергетик', 'писта', 'pepsi', 'sigaret'] },
  { id: 'food', name: 'Озиқ-овқат', color: '#8b5cf6', tags: ['Нон', 'ун', 'гўшт', 'туз', 'шакар', 'салат', 'ovqat', 'go\'sht'] },
  { id: 'utilities', name: 'Коммунал ва тўловлар', color: '#06b6d4', tags: ['свет', 'газ', 'сув', 'чиқинди', 'svet', 'gaz', 'suv'] },
  { id: 'children', name: 'Болаларга', color: '#f43f5e', tags: ['мактабга буфет пули', 'мактабга йўл пули', 'maktab', 'bolalar'] },
  { id: 'transport', name: 'Транспорт', color: '#3b82f6', tags: ['такси', 'автобус', 'маршрутка', 'дамас', 'taksi', 'avtobus'] },
  { id: 'business', name: 'Бизнес учун', color: '#eab308', tags: ['biznes'] },
  { id: 'communication', name: 'Тўловлар ва Алоқа', color: '#6366f1', tags: ['пайнет', 'вай фай', 'paynet', 'wifi', 'internet'] },
  { id: 'household', name: 'Хўжалик товарлари', color: '#f97316', tags: ['ro\'zg\'or'] },
  { id: 'clothes', name: 'Кийим-кечак', color: '#ec4899', tags: ['kiyim'] },
  { id: 'health', name: 'Соғлиқ', color: '#ef4444', tags: ['дори', 'врач', 'dori', 'kasalxona'] },
];

// Helper transliteration (Latin to Cyrillic for Uzb)
const latinToCyrillic = (str) => {
  if (!str) return '';
  const map = {
    'Sh': 'Ш', 'Ch': 'Ч', 'Ya': 'Я', 'Yu': 'Ю', 'Yo': 'Ё', 'O\'': 'Ў', 'G\'': 'Ғ',
    'sh': 'ш', 'ch': 'ч', 'ya': 'я', 'yu': 'ю', 'yo': 'ё', 'o\'': 'ў', 'g\'': 'ғ',
    'A': 'А', 'B': 'Б', 'D': 'Д', 'E': 'Э', 'F': 'Ф', 'G': 'Г', 'H': 'Ҳ', 'I': 'И',
    'J': 'Ж', 'K': 'К', 'L': 'Л', 'M': 'М', 'N': 'Н', 'O': 'О', 'P': 'П', 'Q': 'Қ',
    'R': 'Р', 'S': 'С', 'T': 'Т', 'U': 'У', 'V': 'В', 'X': 'Х', 'Y': 'Й', 'Z': 'З',
    'a': 'а', 'b': 'б', 'd': 'д', 'e': 'э', 'f': 'ф', 'g': 'г', 'h': 'ҳ', 'i': 'и',
    'j': 'ж', 'k': 'к', 'l': 'л', 'm': 'м', 'n': 'н', 'o': 'о', 'p': 'п', 'q': 'қ',
    'r': 'р', 's': 'с', 't': 'т', 'u': 'у', 'v': 'в', 'x': 'х', 'y': 'й', 'z': 'з'
  };
  let res = str;
  for (let key in map) {
    res = res.split(key).join(map[key]);
  }
  return res;
};

const EXTRA_COLORS = ['#f43f5e', '#ec4899', '#d946ef', '#8b5cf6', '#6366f1', '#3b82f6', '#0ea5e9', '#06b6d4', '#14b8a6', '#10b981', '#22c55e', '#84cc16', '#eab308', '#f59e0b', '#f97316'];

const LOCAL_STORAGE_KEY = 'family_budget_data';
const INCOME_STORAGE_KEY = 'family_budget_income';
const CAT_STORAGE_KEY = 'family_budget_categories';

export default function App() {
  const [activeTab, setActiveTab] = useState('dashboard');
  const [expenses, setExpenses] = useState([]);
  const [monthlyIncome, setMonthlyIncome] = useState('');
  const [categories, setCategories] = useState(DEFAULT_CATEGORIES);
  
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [selectedCategoryModal, setSelectedCategoryModal] = useState(null);

  // Form State
  const [amount, setAmount] = useState('');
  const [category, setCategory] = useState('');
  const [note, setNote] = useState('');
  const [date, setDate] = useState(new Date().toISOString().split('T')[0]);

  // Category Manage State
  const [newCatName, setNewCatName] = useState('');

  // Agent State
  const [agentText, setAgentText] = useState('');
  const [isListening, setIsListening] = useState(false);
  const recognitionRef = useRef(null);

  // Load from local storage
  useEffect(() => {
    const data = localStorage.getItem(LOCAL_STORAGE_KEY);
    const income = localStorage.getItem(INCOME_STORAGE_KEY);
    const storedCats = localStorage.getItem(CAT_STORAGE_KEY);
    
    if (data) { try { setExpenses(JSON.parse(data)); } catch (e) { console.error(e); } }
    if (income) setMonthlyIncome(income);
    if (storedCats) {
      try { 
        const parsed = JSON.parse(storedCats);
        if (parsed.length > 0) {
          // Merge missing tags property for backward compatibility
          const updatedCats = parsed.map(c => ({...c, tags: c.tags || []}));
          // Merge missing default categories if user hasn't deleted them
          const currentIds = new Set(updatedCats.map(c => c.id));
          const missingDefaults = DEFAULT_CATEGORIES.filter(dc => !currentIds.has(dc.id) && (dc.id === 'motivation' || dc.id === 'badhabits')); // Force motivation and badhabits
          setCategories([...updatedCats, ...missingDefaults]);
        }
      } catch (e) { console.error(e); }
    }
  }, []);

  useEffect(() => {
    if (!category && categories.length > 1) {
      setCategory(categories.find(c => c.id === 'food')?.id || categories[0].id);
    }
  }, [categories, category]);

  useEffect(() => {
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(expenses));
    localStorage.setItem(INCOME_STORAGE_KEY, monthlyIncome);
    localStorage.setItem(CAT_STORAGE_KEY, JSON.stringify(categories));
  }, [expenses, monthlyIncome, categories]);

  const handleAddExpense = (e) => {
    e.preventDefault();
    if (!amount || isNaN(amount)) return;

    const newExpense = {
      id: uuidv4(),
      amount: parseFloat(amount),
      category,
      note,
      date,
      timestamp: Date.now()
    };

    setExpenses(prev => [...prev, newExpense]);
    setAmount('');
    setNote('');
    setIsAddModalOpen(false);
  };

  const handleDelete = (id) => {
    setExpenses(prev => prev.filter(e => e.id !== id));
  };

  const handleAddCategory = () => {
    if (!newCatName.trim()) return;
    const newCat = {
      id: uuidv4(),
      name: newCatName.trim(),
      color: EXTRA_COLORS[categories.length % EXTRA_COLORS.length],
      tags: []
    };
    setCategories(prev => [...prev, newCat]);
    setNewCatName('');
  };

  const handleDeleteCategory = (catId) => {
    if (catId === 'savings' || catId === 'motivation' || catId === 'badhabits') {
      alert("Бу махсус категорияни ўчириш мумкин эмас.");
      return;
    }
    if (window.confirm('Ушбу категорияни ўчирмоқчимисиз? (Унга тегишли харажатлар ўчмайди, лекин номи йўқолади)')) {
      setCategories(prev => prev.filter(c => c.id !== catId));
    }
  };

  const handleEditCategory = (catId) => {
    const cat = categories.find(c => c.id === catId);
    if (!cat) return;
    const newName = prompt('Янги номни киритинг:', cat.name);
    if (newName && newName.trim()) {
      setCategories(prev => prev.map(c => c.id === catId ? { ...c, name: latinToCyrillic(newName.trim()) } : c));
    }
  };

  // Agent Magic Parsing
  const handleAgentSubmit = () => {
    if(!agentText.trim()) return;
    let txt = agentText.toLowerCase();
    
    // Replace string numbers before extracting
    txt = txt.replace(/ming/g, '000').replace(/минг/g, '000');
    txt = txt.replace(/mln/g, '000000').replace(/млн/g, '000000').replace(/million/g, '000000').replace(/миллион/g, '000000');
    
    // Convert e.g. "2 000 000" into "2000000" for precise match
    const cleanNumTxt = txt.replace(/(\d)\s+(?=\d)/g, '$1');
    const numberMatches = cleanNumTxt.match(/\d+/g);
    
    let amt = 0;
    if(numberMatches && numberMatches.length > 0) {
       amt = parseFloat(numberMatches[numberMatches.length - 1]);
    }
    
    // Clean up text
    let nText = txt.replace(/\d+/g, '').replace(/(харажат|xarajat|сўм|sum|ming|минг|mln|млн)/gi, '').replace(/\s+/g, ' ').trim();
    if (!nText) nText = 'Агент орқали';

    // Transliterate to Cyrillic for nicer display
    const cyrillicNote = latinToCyrillic(nText);

    let foundCatId = 'food';
    let matched = false;
    
    for(const cat of categories) {
      if(cat.name.toLowerCase().includes(nText) || cat.name.toLowerCase().includes(cyrillicNote)) {
        foundCatId = cat.id; matched = true; break;
      }
      for(const tag of (cat.tags || [])) {
        if (nText.includes(tag.toLowerCase()) || cyrillicNote.includes(tag.toLowerCase()) || tag.toLowerCase().includes(nText)) {
          foundCatId = cat.id; matched = true; break;
        }
      }
      if(matched) break;
    }

    const newExp = {
      id: uuidv4(),
      amount: amt > 0 ? amt : 0,
      category: foundCatId,
      note: cyrillicNote.charAt(0).toUpperCase() + cyrillicNote.slice(1),
      date: new Date().toISOString().split('T')[0],
      timestamp: Date.now()
    };
    
    setExpenses(prev => [...prev, newExp]);
    setAgentText('');
    alert(`🤖 Агент: "${newExp.note}" ${amt > 0 ? amt + ' сўм' : '(суммасиз)'} қўшилди!`);
  };

  const handleVoice = () => {
    if (isListening) {
      if (recognitionRef.current) recognitionRef.current.stop();
      setIsListening(false);
      return;
    }

    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
      alert("Браузерингиз овозли киритишни қўллаб-қувватламайди.");
      return;
    }

    const recognition = new SpeechRecognition();
    recognitionRef.current = recognition;
    recognition.lang = 'uz-UZ';
    recognition.continuous = true;
    recognition.interimResults = true;
    
    let finalTranscript = '';

    recognition.onstart = () => setIsListening(true);
    recognition.onresult = (event) => {
      let interim = '';
      for (let i = event.resultIndex; i < event.results.length; ++i) {
        if (event.results[i].isFinal) {
          finalTranscript += event.results[i][0].transcript + ' ';
        } else {
          interim += event.results[i][0].transcript;
        }
      }
      setAgentText(latinToCyrillic(finalTranscript + interim));
      
      // Auto stop after 3 seconds of silence
      clearTimeout(window.voiceTimeout);
      window.voiceTimeout = setTimeout(() => {
         recognition.stop();
         setIsListening(false);
      }, 3000);
    };
    
    recognition.onerror = () => setIsListening(false);
    recognition.onend = () => setIsListening(false);
    
    setAgentText('');
    recognition.start();
  };

  // Sync
  const handleExport = () => {
    const dataStr = JSON.stringify({ expenses, categories });
    navigator.clipboard.writeText(dataStr);
    alert('Маълумотлар нусхаланди! Энди уни Telegram орқали юборинг.');
  };

  const handleImport = () => {
    const dataStr = prompt('Келиб тушган маълумотларни (кодни) шу ерга жойланг:');
    if (dataStr) {
      try {
        const importedData = JSON.parse(dataStr);
        if (importedData.expenses) {
          const currentIds = new Set(expenses.map(e => e.id));
          const newItems = importedData.expenses.filter(e => !currentIds.has(e.id));
          setExpenses(prev => [...prev, ...newItems]);
        }
        if (importedData.categories) {
           const currentCatIds = new Set(categories.map(c => c.id));
           const newCats = importedData.categories.filter(c => !currentCatIds.has(c.id));
           if (newCats.length > 0) {
              setCategories(prev => [...prev, ...newCats.map(c => ({...c, tags: c.tags || []}))]);
           }
        }
        alert(`Маълумотлар муваффақиятли қабул қилинди!`);
      } catch (e) {
        alert('Нотўғри форматдаги маълумот.');
      }
    }
  };

  // Excel Export Magic
  const handleExcelExport = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Ойлик Хисобот');
    
    const cols = [
      { header: 'Сана', key: 'date', width: 12 },
      { header: 'Хафта куни', key: 'weekday', width: 15 },
      ...categories.map(c => ({ header: c.name, key: c.id, width: 20 })),
      { header: 'Кунлик Жами', key: 'dailyTotal', width: 20 }
    ];
    worksheet.columns = cols;
    
    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };

    const now = new Date();
    const days = eachDayOfInterval({ start: startOfMonth(now), end: endOfMonth(now) });
    
    let totalAll = 0;
    const catTotals = {};
    categories.forEach(c => catTotals[c.id] = 0);

    const uzbekDays = ['Якшанба', 'Душанба', 'Сешанба', 'Чоршанба', 'Пайшанба', 'Жума', 'Шанба'];

    days.forEach(day => {
       const dateStr = format(day, 'yyyy-MM-dd');
       const isSun = isSunday(day);
       const weekdayName = uzbekDays[day.getDay()];
       
       const rowData = { date: format(day, 'dd.MM.yyyy'), weekday: weekdayName };
       let dailySum = 0;

       categories.forEach(cat => {
          const catSum = expenses.filter(e => e.date === dateStr && e.category === cat.id).reduce((sum, e) => sum + e.amount, 0);
          rowData[cat.id] = catSum > 0 ? catSum : '';
          if (cat.id !== 'savings' && cat.id !== 'motivation') {
             dailySum += catSum;
          }
          catTotals[cat.id] += catSum;
       });
       
       rowData.dailyTotal = dailySum > 0 ? dailySum : '';
       totalAll += dailySum;

       const row = worksheet.addRow(rowData);
       
       if (isSun) {
         row.eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
         });
       }
    });

    const totalRowData = { date: 'ЖАМИ:', weekday: '' };
    categories.forEach(cat => {
       totalRowData[cat.id] = catTotals[cat.id] > 0 ? catTotals[cat.id] : 0;
    });
    totalRowData.dailyTotal = totalAll;
    
    const tRow = worksheet.addRow(totalRowData);
    tRow.font = { bold: true };
    tRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE047' } };

    worksheet.eachRow((row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `Бюджет_${format(now, 'MM_yyyy')}.xlsx`);
  };

  const formatMoney = (val) => {
    if (!val) return '0 сўм';
    return new Intl.NumberFormat('ru-RU').format(val) + ' сўм';
  };
  
  const formatCompact = (val) => {
    if (val === 0) return '';
    return new Intl.NumberFormat('ru-RU').format(val);
  };

  const currentMonthExpenses = useMemo(() => {
    const now = new Date();
    return expenses.filter(e => isSameMonth(parseISO(e.date), now));
  }, [expenses]);

  const totalExpense = currentMonthExpenses
    .filter(e => e.category !== 'savings' && e.category !== 'motivation')
    .reduce((sum, e) => sum + e.amount, 0);

  const totalSavings = currentMonthExpenses
    .filter(e => e.category === 'savings')
    .reduce((sum, e) => sum + e.amount, 0);

  const totalMotivation = currentMonthExpenses
    .filter(e => e.category === 'motivation')
    .reduce((sum, e) => sum + e.amount, 0);

  const remainingBalance = (parseFloat(monthlyIncome) || 0) - totalExpense - totalSavings;

  const categoryData = useMemo(() => {
    const data = categories.map(cat => ({
      id: cat.id,
      name: cat.name,
      value: currentMonthExpenses.filter(e => e.category === cat.id).reduce((sum, e) => sum + e.amount, 0),
      color: cat.color
    })).filter(d => d.value > 0);
    return data.sort((a, b) => b.value - a.value);
  }, [currentMonthExpenses, categories]);

  const chartData = useMemo(() => {
    const now = new Date();
    const days = eachDayOfInterval({ start: startOfMonth(now), end: endOfMonth(now) });
    
    const currentDay = parseInt(format(now, 'dd'), 10);
    const data = days.map(day => {
      const dateStr = format(day, 'yyyy-MM-dd');
      const dayExpenses = currentMonthExpenses.filter(e => e.date === dateStr && e.category !== 'savings' && e.category !== 'motivation');
      const dailySum = dayExpenses.reduce((sum, e) => sum + e.amount, 0);
      return {
        name: format(day, 'dd'),
        sum: dailySum > 0 ? dailySum : null,
      };
    });
    
    return data.filter((d, i) => i < currentDay);
  }, [currentMonthExpenses]);

  const groupedExpenses = useMemo(() => {
    const groups = {};
    [...expenses].sort((a,b) => new Date(b.date) - new Date(a.date) || b.timestamp - a.timestamp).forEach(exp => {
      if (!groups[exp.date]) groups[exp.date] = [];
      groups[exp.date].push(exp);
    });
    return groups;
  }, [expenses]);

  const renderCustomizedLabel = (props) => {
    const { cx, cy, innerRadius, outerRadius, midAngle, percent, fill, value } = props;
    const RADIAN = Math.PI / 180;
    // Put label nicely inside the slice
    const radius = innerRadius + (outerRadius - innerRadius) * 0.5;
    const x = cx + radius * Math.cos(-midAngle * RADIAN);
    const y = cy + radius * Math.sin(-midAngle * RADIAN);
  
    // Only show if the slice is large enough
    if (percent < 0.05) return null;
  
    return (
      <text 
        x={x} 
        y={y} 
        fill="#ffffff" 
        textAnchor="middle" 
        dominantBaseline="central" 
        fontSize={11} 
        fontWeight="bold"
        style={{ textShadow: '0px 1px 2px rgba(0,0,0,0.5)' }}
      >
        {formatCompact(value)}
      </text>
    );
  };

  const selectedCatObj = categories.find(c => c.id === category);
  const noteSuggestions = useMemo(() => {
    if(!selectedCatObj) return [];
    const historyNotes = expenses.filter(e => e.category === selectedCatObj.id && e.note).map(e => e.note.trim());
    const combined = [...(selectedCatObj.tags || []), ...historyNotes];
    return [...new Set(combined)];
  }, [category, expenses, selectedCatObj]);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 pb-20 font-sans selection:bg-indigo-200">
      <header className="bg-white shadow-sm sticky top-0 z-50">
        <div className="max-w-3xl mx-auto px-4 py-4 flex justify-between items-center">
          <div className="flex items-center gap-2">
            <div className="bg-indigo-600 p-2 rounded-lg text-white shadow-md">
              <Wallet size={24} />
            </div>
            <h1 className="text-xl font-bold bg-gradient-to-r from-indigo-600 to-violet-600 bg-clip-text text-transparent">
              Оилавий Бюджет
            </h1>
          </div>
          <button 
            onClick={() => setIsAddModalOpen(true)}
            className="bg-indigo-600 hover:bg-indigo-700 text-white p-2 rounded-full shadow-lg transition-transform hover:scale-105 active:scale-95"
          >
            <Plus size={24} />
          </button>
        </div>
      </header>

      <main className="max-w-3xl mx-auto px-4 py-6 space-y-6">
        
        {activeTab === 'dashboard' && (
          <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            
            {/* Agent Bar */}
            <div className="bg-white p-2 pl-4 rounded-full shadow-sm border border-slate-100 flex items-center gap-2 focus-within:ring-2 ring-indigo-500">
               <Bot size={20} className="text-indigo-400" />
               <input 
                 type="text" 
                 value={agentText}
                 onChange={(e) => setAgentText(latinToCyrillic(e.target.value))}
                 onKeyDown={(e) => e.key === 'Enter' && handleAgentSubmit()}
                 placeholder="Агентга айтинг: 'Нон 5000' ёки 'Такси 20000'"
                 className="flex-1 bg-transparent outline-none text-sm font-medium"
               />
               <button onClick={handleVoice} className={`p-2 rounded-full transition-colors ${isListening ? 'bg-red-100 text-red-500 animate-pulse' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>
                 {isListening ? <div className="w-4 h-4 rounded-sm bg-red-500"></div> : <Mic size={18} />}
               </button>
               <button onClick={handleAgentSubmit} className="bg-indigo-600 text-white p-2 px-4 rounded-full text-sm font-bold shadow-md hover:bg-indigo-700 active:scale-95 transition-all">
                 Қўшиш
               </button>
            </div>

            <div className="bg-white rounded-2xl p-5 shadow-sm border border-slate-100 relative overflow-hidden group">
              <div className="absolute top-0 right-0 w-32 h-32 bg-indigo-50 rounded-full -mr-16 -mt-16 transition-transform group-hover:scale-110"></div>
              
              <div className="relative z-10">
                <label className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Жорий Ойлик Маош</label>
                <div className="mt-1 mb-4 flex items-center border-b border-slate-200 pb-2 focus-within:border-indigo-500 transition-colors">
                  <input 
                    type="number" 
                    value={monthlyIncome}
                    onChange={(e) => setMonthlyIncome(e.target.value)}
                    placeholder="5000000"
                    className="w-full text-2xl font-bold text-slate-800 outline-none bg-transparent"
                  />
                  <span className="text-slate-400 font-medium ml-2">сўм</span>
                </div>

                <div className="flex justify-between items-end mt-4">
                  <div>
                    <div className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">Қолдиқ</div>
                    <div className={`text-2xl font-bold ${remainingBalance < 0 ? 'text-red-500' : 'text-emerald-600'}`}>
                      {formatMoney(remainingBalance)}
                    </div>
                  </div>
                  <div className="text-right">
                    <div className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">Запас/Сармоя</div>
                    <div className="text-xl font-bold text-slate-800">
                      {formatMoney(totalSavings)}
                    </div>
                  </div>
                </div>

                {/* Motivation Info */}
                <div className="mt-4 pt-4 border-t border-slate-100 flex items-center justify-between">
                   <div className="flex items-center gap-2 text-emerald-600">
                     <Target size={18} />
                     <span className="text-sm font-bold uppercase tracking-wide">Мотивация заҳираси:</span>
                   </div>
                   <span className="font-bold text-emerald-600 bg-emerald-50 px-2 py-1 rounded-lg">
                      {formatMoney(totalMotivation)}
                   </span>
                </div>

              </div>
            </div>

            <div className="bg-white rounded-2xl p-5 shadow-sm border border-slate-100">
              <h3 className="text-lg font-semibold mb-2 text-slate-800">Категориялар бўйича</h3>
              {categoryData.length > 0 ? (
                <>
                  <div className="h-[250px] w-full relative">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie
                          data={categoryData}
                          cx="50%"
                          cy="50%"
                          innerRadius={65}
                          outerRadius={105}
                          paddingAngle={2}
                          dataKey="value"
                          labelLine={false}
                          label={renderCustomizedLabel}
                        >
                          {categoryData.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={entry.color} />
                          ))}
                        </Pie>
                        <RechartsTooltip 
                          formatter={(value) => formatMoney(value)}
                          contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                        />
                      </PieChart>
                    </ResponsiveContainer>
                    <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                      <span className="text-[10px] text-slate-400 uppercase font-bold tracking-wider">Жами Харажат</span>
                      <span className="text-sm font-black text-slate-800">{formatCompact(totalExpense)}</span>
                    </div>
                  </div>

                  <div className="mt-4 space-y-2">
                    {categoryData.map(cat => (
                      <button 
                        key={cat.id}
                        onClick={() => setSelectedCategoryModal(cat.id)}
                        className="w-full flex items-center justify-between p-3 rounded-xl hover:bg-slate-50 transition-colors border border-transparent hover:border-slate-100 group"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-3 h-3 rounded-full" style={{ backgroundColor: cat.color }}></div>
                          <span className="font-medium text-slate-700">{cat.name}</span>
                        </div>
                        <div className="flex items-center gap-2">
                          <span className="font-bold text-slate-900">{formatMoney(cat.value)}</span>
                          <ChevronRight size={16} className="text-slate-300 group-hover:text-indigo-400 transition-colors" />
                        </div>
                      </button>
                    ))}
                  </div>
                </>
              ) : (
                <div className="h-[200px] flex items-center justify-center text-slate-400 text-sm">
                  Харажатлар йўқ
                </div>
              )}
            </div>

            <div className="bg-white rounded-2xl p-5 shadow-sm border border-slate-100">
              <h3 className="text-lg font-semibold mb-4 text-slate-800">Кунлик харажатлар</h3>
              <div className="h-[240px] w-full mt-4">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={chartData} margin={{ top: 20, right: 0, left: -20, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#94a3b8' }} />
                    <RechartsTooltip 
                      cursor={{ fill: '#f8fafc' }}
                      formatter={(value) => formatMoney(value)}
                      contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                      labelStyle={{ color: '#64748b', marginBottom: '4px' }}
                    />
                    <Bar 
                      dataKey="sum" 
                      fill="#4f46e5" 
                      radius={[4, 4, 0, 0]} 
                      barSize={20}
                    >
                      <LabelList dataKey="sum" position="top" formatter={formatCompact} style={{ fontSize: '8px', fill: '#64748b', fontWeight: 600 }} offset={5} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
        )}

        {activeTab === 'list' && (
          <div className="space-y-4 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <h2 className="text-xl font-bold text-slate-800 px-1">Барча харажатлар</h2>
            
            {Object.keys(groupedExpenses).length === 0 ? (
              <div className="text-center text-slate-400 py-10">Маълумот йўқ</div>
            ) : (
              <div className="space-y-6">
                {Object.entries(groupedExpenses).map(([dateStr, items]) => (
                  <div key={dateStr} className="space-y-2">
                    <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider px-2">{dateStr}</h3>
                    <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
                      {items.map((exp, idx) => {
                        const cat = categories.find(c => c.id === exp.category) || { name: 'Ўчирилган категория', color: '#ccc' };
                        return (
                          <div key={exp.id} className={`p-4 flex items-center justify-between group ${idx !== items.length - 1 ? 'border-b border-slate-100' : ''}`}>
                            <div className="flex items-center gap-3">
                              <div className="w-1.5 h-10 rounded-full" style={{ backgroundColor: cat.color }}></div>
                              <div>
                                <p className="font-semibold text-slate-800 text-sm">{cat.name}</p>
                                {exp.note && <p className="text-xs text-slate-500 mt-0.5">{exp.note}</p>}
                              </div>
                            </div>
                            <div className="text-right flex flex-col items-end">
                              <p className={`font-bold text-sm ${exp.category === 'savings' || exp.category === 'motivation' ? 'text-emerald-500' : exp.category === 'badhabits' ? 'text-red-500' : 'text-slate-800'}`}>
                                {exp.category === 'savings' || exp.category === 'motivation' ? '+' : '-'}{formatMoney(exp.amount)}
                              </p>
                              <button 
                                onClick={() => handleDelete(exp.id)}
                                className="text-red-400 hover:text-red-600 opacity-0 group-hover:opacity-100 transition-opacity mt-1"
                              >
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {activeTab === 'settings' && (
          <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
            
            <div>
              <h2 className="text-xl font-bold text-slate-800 px-1 mb-4">Маълумотларни бошқариш</h2>
              <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-100 space-y-6">
                
                <button 
                    onClick={handleExcelExport}
                    className="flex items-center justify-center gap-3 w-full bg-emerald-50 text-emerald-700 p-4 rounded-xl font-bold hover:bg-emerald-100 transition-colors shadow-sm"
                  >
                    <Download size={20} />
                    Excel (Жадвал) га юклаб олиш
                </button>

                <p className="text-sm text-slate-600 leading-relaxed pt-2 border-t border-slate-100">
                  Бу маълумотлар фақат қурилмангизда сақланади. Оила аъзонгиз билан базани бирлаштириш учун нусхалаб юборинг.
                </p>
                
                <div className="grid gap-4">
                  <button 
                    onClick={handleExport}
                    className="flex items-center justify-center gap-3 w-full bg-indigo-50 text-indigo-700 p-4 rounded-xl font-bold hover:bg-indigo-100 transition-colors"
                  >
                    <ArrowUpFromLine size={20} />
                    Маълумотларни нусхалаш
                  </button>
                  <button 
                    onClick={handleImport}
                    className="flex items-center justify-center gap-3 w-full bg-slate-50 text-slate-700 border border-slate-200 p-4 rounded-xl font-bold hover:bg-slate-100 transition-colors"
                  >
                    <ArrowDownToLine size={20} />
                    Кодни қабул қилиш
                  </button>
                </div>
              </div>
            </div>

            <div>
              <h2 className="text-xl font-bold text-slate-800 px-1 mb-4">Категорияларни созлаш</h2>
              <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-100 space-y-4">
                
                <div className="flex gap-2">
                  <input 
                    type="text" 
                    value={newCatName}
                    onChange={(e) => setNewCatName(latinToCyrillic(e.target.value))}
                    placeholder="Янги категория номи..."
                    className="flex-1 border border-slate-200 rounded-xl p-3 text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
                  />
                  <button 
                    onClick={handleAddCategory}
                    className="bg-indigo-600 text-white px-4 rounded-xl font-bold hover:bg-indigo-700 transition-colors"
                  >
                    Қўшиш
                  </button>
                </div>

                <div className="space-y-2 mt-4 max-h-80 overflow-y-auto pr-2">
                  {categories.map(cat => (
                    <div key={cat.id} className="flex justify-between items-center p-3 border border-slate-100 rounded-xl hover:bg-slate-50">
                      <div className="flex items-center gap-3">
                        <div className="w-3 h-3 rounded-full" style={{ backgroundColor: cat.color }}></div>
                        <span className="font-medium text-slate-700 text-sm">{cat.name}</span>
                      </div>
                      <div className="flex gap-2">
                        {cat.id !== 'savings' && cat.id !== 'motivation' && cat.id !== 'badhabits' && (
                          <button onClick={() => handleEditCategory(cat.id)} className="text-slate-400 hover:text-indigo-500 p-1 transition-colors">
                            <Edit2 size={16} />
                          </button>
                        )}
                        {cat.id !== 'savings' && cat.id !== 'motivation' && cat.id !== 'badhabits' && (
                          <button onClick={() => handleDeleteCategory(cat.id)} className="text-slate-400 hover:text-red-500 p-1 transition-colors">
                            <Trash2 size={16} />
                          </button>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
                
              </div>
            </div>
            
          </div>
        )}

      </main>

      <nav className="fixed bottom-0 w-full bg-white border-t border-slate-200 pb-safe z-40">
        <div className="max-w-md mx-auto flex justify-around">
          <button 
            onClick={() => setActiveTab('dashboard')} 
            className={`p-4 flex flex-col items-center gap-1 w-full transition-colors ${activeTab === 'dashboard' ? 'text-indigo-600' : 'text-slate-400 hover:text-slate-600'}`}
          >
            <Home size={20} className={activeTab === 'dashboard' ? 'fill-indigo-100' : ''} />
            <span className="text-[10px] font-bold">Асосий</span>
          </button>
          <button 
            onClick={() => setActiveTab('list')} 
            className={`p-4 flex flex-col items-center gap-1 w-full transition-colors ${activeTab === 'list' ? 'text-indigo-600' : 'text-slate-400 hover:text-slate-600'}`}
          >
            <ChartIcon size={20} className={activeTab === 'list' ? 'fill-indigo-100' : ''} />
            <span className="text-[10px] font-bold">Рўйхат</span>
          </button>
          <button 
            onClick={() => setActiveTab('settings')} 
            className={`p-4 flex flex-col items-center gap-1 w-full transition-colors ${activeTab === 'settings' ? 'text-indigo-600' : 'text-slate-400 hover:text-slate-600'}`}
          >
            <Settings size={20} className={activeTab === 'settings' ? 'fill-indigo-100' : ''} />
            <span className="text-[10px] font-bold">Созлама</span>
          </button>
        </div>
      </nav>

      {selectedCategoryModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex justify-center items-end sm:items-center p-4">
          <div className="bg-slate-50 w-full max-w-md rounded-3xl shadow-2xl h-[80vh] flex flex-col animate-in slide-in-from-bottom-full duration-300 overflow-hidden">
            <div className="bg-white p-5 border-b border-slate-100 flex justify-between items-center shrink-0">
              <div className="flex items-center gap-3">
                <div 
                  className="w-4 h-4 rounded-full shadow-sm" 
                  style={{ backgroundColor: categories.find(c => c.id === selectedCategoryModal)?.color }}
                ></div>
                <h2 className="text-xl font-bold text-slate-800">
                  {categories.find(c => c.id === selectedCategoryModal)?.name}
                </h2>
              </div>
              <button onClick={() => setSelectedCategoryModal(null)} className="bg-slate-100 text-slate-500 hover:bg-slate-200 p-2 rounded-full transition-colors">
                <X size={20} />
              </button>
            </div>
            
            <div className="p-4 overflow-y-auto flex-1">
              <div className="space-y-3">
                {expenses
                  .filter(e => e.category === selectedCategoryModal)
                  .sort((a,b) => new Date(b.date) - new Date(a.date))
                  .map(exp => (
                    <div key={exp.id} className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 flex justify-between items-center">
                      <div>
                        <p className="text-sm font-bold text-slate-800">{exp.date}</p>
                        {exp.note && <p className="text-xs text-slate-500 mt-1">{exp.note}</p>}
                      </div>
                      <p className="font-bold text-slate-900 text-lg">{formatMoney(exp.amount)}</p>
                    </div>
                ))}
                {expenses.filter(e => e.category === selectedCategoryModal).length === 0 && (
                   <div className="text-center text-slate-400 py-10 font-medium">Бу категорияда ҳали харажат йўқ</div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {isAddModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex justify-center items-end sm:items-center">
          <div className="bg-white w-full max-w-md rounded-t-3xl sm:rounded-3xl p-6 shadow-2xl animate-in slide-in-from-bottom-full sm:slide-in-from-bottom-0 sm:zoom-in-95 duration-300">
            <div className="flex justify-between items-center mb-6">
              <h2 className="text-2xl font-bold text-slate-800">Янги ёзув</h2>
              <button onClick={() => setIsAddModalOpen(false)} className="bg-slate-100 text-slate-500 hover:bg-slate-200 p-2 rounded-full transition-colors">
                <X size={20} />
              </button>
            </div>
            
            <form onSubmit={handleAddExpense} className="space-y-5">
              <div>
                <label className="block text-sm font-bold text-slate-700 mb-2">Сумма (сўм)</label>
                <div className="relative">
                  <input 
                    type="number" 
                    value={amount}
                    onChange={(e) => setAmount(e.target.value)}
                    placeholder="150000"
                    className="w-full border-2 border-slate-200 rounded-2xl p-4 text-xl font-bold text-slate-800 focus:ring-0 focus:border-indigo-500 outline-none transition-colors"
                    required
                  />
                </div>
              </div>
              
              <div>
                <label className="block text-sm font-bold text-slate-700 mb-2">Категория</label>
                <div className="flex flex-wrap gap-2 mt-2 max-h-32 overflow-y-auto p-1">
                  {categories.map(cat => (
                    <button
                      key={cat.id}
                      type="button"
                      onClick={() => setCategory(cat.id)}
                      className={`px-4 py-2 rounded-xl text-xs font-bold border-2 transition-all ${
                        category === cat.id 
                        ? 'border-indigo-500 bg-indigo-50 text-indigo-700 shadow-sm scale-105' 
                        : 'border-slate-100 bg-white text-slate-600 hover:border-slate-300 hover:bg-slate-50'
                      }`}
                    >
                      {cat.name}
                    </button>
                  ))}
                </div>
              </div>

              <div>
                <label className="block text-sm font-bold text-slate-700 mb-2">Нима учун? (Авто-таклиф)</label>
                <input 
                  type="text" 
                  list="category-tags"
                  value={note}
                  onChange={(e) => setNote(latinToCyrillic(e.target.value))}
                  placeholder="Масалан: нон, электр, мактаб..."
                  className="w-full border-2 border-slate-200 rounded-2xl p-3 text-base focus:border-indigo-500 outline-none transition-colors"
                />
                <datalist id="category-tags">
                  {noteSuggestions.map((s,i) => <option key={i} value={s} />)}
                </datalist>
              </div>

              <div>
                <label className="block text-sm font-bold text-slate-700 mb-2">Сана</label>
                <input 
                  type="date" 
                  value={date}
                  onChange={(e) => setDate(e.target.value)}
                  className="w-full border-2 border-slate-200 rounded-2xl p-3 text-base focus:border-indigo-500 outline-none transition-colors"
                />
              </div>
              
              <div className="pt-2">
                <button 
                  type="submit"
                  className="w-full bg-indigo-600 text-white font-bold text-lg rounded-2xl p-4 shadow-lg shadow-indigo-200 hover:bg-indigo-700 hover:shadow-indigo-300 transition-all active:scale-[0.98]"
                >
                  Қўшиш
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

    </div>
  );
}

