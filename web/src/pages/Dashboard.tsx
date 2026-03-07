import { useState, useRef, useEffect, useCallback } from 'react';
import { useQuery } from '@tanstack/react-query';
import { useLocation } from 'wouter';
import {
    Mail, Trash2, Sparkles,
    Search, AlertTriangle, BarChart3, RefreshCw, LogOut, Loader2,
    Paperclip, Download, ChevronDown, ChevronUp, Check, X, Tag, Info, Send,
    Menu, ChevronLeft, ChevronRight, Edit2
} from 'lucide-react';

const API_BASE = 'http://localhost:8000';

export default function Dashboard() {
    const [, setLocation] = useLocation();
    const [activeTab, setActiveTab] = useState<'inbox' | 'sent' | 'cleanup' | 'stats' | string>('inbox');
    const [searchQuery, setSearchQuery] = useState('');
    const [selectedIds, setSelectedIds] = useState<string[]>([]);
    const [limit, setLimit] = useState(20);
    const [selectedEmail, setSelectedEmail] = useState<any>(null);
    const [isSidebarOpen, setIsSidebarOpen] = useState(true);
    const [isListOpen, setIsListOpen] = useState(true);
    const [isSummarizing, setIsSummarizing] = useState(false);
    const [isCategorizing, setIsCategorizing] = useState(false);
    const [isBatchCategorizing, setIsBatchCategorizing] = useState(false);
    const [expandedEmails, setExpandedEmails] = useState<Record<string, boolean>>({});
    const [isFetchingMore, setIsFetchingMore] = useState(false);
    const [isComposeOpen, setIsComposeOpen] = useState(false);
    const [composeForm, setComposeForm] = useState({ to: '', subject: '', body: '' });
    const [isSending, setIsSending] = useState(false);
    const listRef = useRef<HTMLDivElement>(null);

    const token = localStorage.getItem('auth_token');

    // Handle logout
    const handleLogout = () => {
        localStorage.removeItem('auth_token');
        setLocation('/login');
    };

    // Setup authorized fetch
    const authorizedFetch = async (url: string, options: any = {}) => {
        const headers = {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
            ...options.headers,
        };

        const res = await fetch(url, { ...options, headers });
        if (res.status === 401) {
            handleLogout();
            throw new Error("Unauthorized");
        }
        return res;
    };

    // Fetch emails
    const { data: emails, isLoading, refetch } = useQuery({
        queryKey: ['emails', searchQuery, limit, activeTab], // Added activeTab to queryKey for re-fetching on tab change
        queryFn: async () => {
            let endpoint = '';
            if (searchQuery) {
                // Scope search based on active tab
                const scopedQuery = activeTab === 'sent' ? `in:sent ${searchQuery}` : `in:inbox ${searchQuery}`;
                endpoint = `${API_BASE}/emails/search/${encodeURIComponent(scopedQuery)}?limit=${limit}`;
            } else if (activeTab === 'sent') {
                endpoint = `${API_BASE}/emails/sent?limit=${limit}`;
            } else {
                endpoint = `${API_BASE}/emails/search/in:inbox?limit=${limit}`;
            }
            const res = await authorizedFetch(endpoint);
            const data = await res.json();
            return data;
        },
        enabled: !!token,
        retry: 1
    });

    // Fetch cleanup suggestions
    const { data: cleanupData, isLoading: isLoadingCleanup } = useQuery({
        queryKey: ['cleanup'],
        queryFn: async () => {
            const res = await authorizedFetch(`${API_BASE}/emails/cleanup-suggestions`);
            return res.json();
        },
        enabled: !!token && activeTab === 'cleanup',
        retry: 1
    });

    const handleSummarize = async (id: string) => {
        if (isSummarizing) return;
        setIsSummarizing(true);
        try {
            const res = await authorizedFetch(`${API_BASE}/emails/${id}/summarize`, { method: 'POST' });
            const data = await res.json();
            setSelectedEmail((prev: any) => ({ ...prev, summary: data.summary }));
        } finally {
            setIsSummarizing(false);
        }
    };

    const handleCategorize = async (id: string, currentEmail: any) => {
        if (isCategorizing) return;
        setIsCategorizing(true);
        try {
            const res = await authorizedFetch(`${API_BASE}/emails/${id}/categorize`, { method: 'POST' });
            const data = await res.json();
            const updated = { ...currentEmail, category: data.category, reasoning: data.reasoning };
            if (selectedEmail?.id === id) setSelectedEmail(updated);

            // Remove from list if it was moved to a label (since we are in 'Smart Inbox' usually viewing INBOX)
            refetch();
        } finally {
            setIsCategorizing(false);
        }
    };

    const handleBatchCategorize = async () => {
        if (isBatchCategorizing || selectedIds.length === 0) return;
        setIsBatchCategorizing(true);
        try {
            await authorizedFetch(`${API_BASE}/emails/batch/categorize`, {
                method: 'POST',
                body: JSON.stringify({ email_ids: selectedIds })
            });
            setSelectedIds([]);
            refetch();
        } finally {
            setIsBatchCategorizing(false);
        }
    };

    const handleRemoveLabel = async (emailId: string, labelName: string) => {
        try {
            await authorizedFetch(`${API_BASE}/emails/${emailId}/labels/${encodeURIComponent(labelName)}`, {
                method: 'DELETE'
            });
            refetch();
            if (selectedEmail?.id === emailId) {
                // Optimistic UI update or just clear to force refresh
                setSelectedEmail(null);
            }
        } catch (err) {
            console.error("Failed to remove label", err);
        }
    };

    const handleSendEmail = async () => {
        if (!composeForm.to || !composeForm.subject || !composeForm.body) return;
        setIsSending(true);
        try {
            await authorizedFetch(`${API_BASE}/emails/send`, {
                method: 'POST',
                body: JSON.stringify(composeForm)
            });
            setIsComposeOpen(false);
            setComposeForm({ to: '', subject: '', body: '' });
            if (activeTab === 'sent') refetch();
        } catch (err) {
            console.error(err);
        } finally {
            setIsSending(false);
        }
    };

    const handleEmailSelect = async (email: any) => {
        setSelectedEmail(email);
        if (email.labels && email.labels.includes('UNREAD')) {
            try {
                await authorizedFetch(`${API_BASE}/emails/${email.id}/read`, { method: 'POST' });
                refetch();
            } catch (err) { }
        }
    };

    const toggleSelect = (id: string, e?: React.MouseEvent) => {
        if (e) e.stopPropagation();
        setSelectedIds(prev =>
            prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
        );
    };

    const toggleSelectAll = () => {
        if (!emails) return;
        if (selectedIds.length === emails.length) {
            setSelectedIds([]);
        } else {
            setSelectedIds(emails.map((e: any) => e.id));
        }
    };

    // Auto-select first email when data loads
    useEffect(() => {
        if (emails && emails.length > 0 && !selectedEmail) {
            setSelectedEmail(emails[0]);
        }
    }, [emails, selectedEmail]);

    // Handle fetching more state
    useEffect(() => {
        setIsFetchingMore(isLoading && limit > 20);
    }, [isLoading, limit]);

    const handleScroll = useCallback(() => {
        if (!listRef.current || isLoading || isFetchingMore) return;

        const { scrollTop, scrollHeight, clientHeight } = listRef.current;
        // If we are within 100px of the bottom
        if (scrollTop + clientHeight >= scrollHeight - 100) {
            if (emails && emails.length >= limit) {
                setLimit(prev => prev + 20);
            }
        }
    }, [isLoading, isFetchingMore, emails, limit]);

    useEffect(() => {
        const listEl = listRef.current;
        if (listEl) {
            listEl.addEventListener('scroll', handleScroll);
            return () => listEl.removeEventListener('scroll', handleScroll);
        }
    }, [handleScroll]);

    useEffect(() => {
        const handleKeyDown = (e: KeyboardEvent) => {
            if (e.key === 'Escape') {
                setSelectedEmail(null);
                setSelectedIds([]);
            }
        };
        window.addEventListener('keydown', handleKeyDown);
        return () => window.removeEventListener('keydown', handleKeyDown);
    }, [setSelectedEmail, setSelectedIds]);

    const getCategoryColor = (cat: string) => {
        const c = cat?.toLowerCase() || '';
        if (c.includes('priority')) return 'bg-blue-100 text-blue-700 border-blue-200';
        if (c.includes('newsletter')) return 'bg-teal-100 text-teal-700 border-teal-200';
        if (c.includes('ad')) return 'bg-amber-100 text-amber-700 border-amber-200';
        if (c.includes('social')) return 'bg-indigo-100 text-indigo-700 border-indigo-200';
        if (c.includes('billing')) return 'bg-emerald-100 text-emerald-700 border-emerald-200';
        if (c.includes('spam')) return 'bg-red-100 text-red-700 border-red-200';
        return 'bg-slate-100 text-slate-700 border-slate-200';
    };

    const formatDate = (dateString: string) => {
        if (!dateString) return '';
        try {
            const date = new Date(dateString);
            if (isNaN(date.getTime())) return dateString;
            return new Intl.DateTimeFormat(undefined, {
                weekday: 'short',
                day: '2-digit',
                month: 'short',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
            }).format(date);
        } catch (e) {
            return dateString;
        }
    };

    return (
        <div className="flex h-screen bg-slate-50 font-sans text-slate-900">
            {/* Sidebar */}
            <aside className={`${isSidebarOpen ? 'w-64' : 'w-20'} bg-white border-r border-slate-200 transition-all duration-300 flex flex-col`}>
                <div className="p-6 flex items-center justify-between">
                    <div className="flex items-center gap-3 overflow-hidden">
                        <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white shrink-0">
                            <Mail size={20} />
                        </div>
                        {isSidebarOpen && <span className="font-bold text-xl tracking-tight animate-in fade-in slide-in-from-left-4 duration-300">MailMaster</span>}
                    </div>
                    <button
                        onClick={() => setIsSidebarOpen(!isSidebarOpen)}
                        className="p-1.5 hover:bg-slate-100 rounded-lg text-slate-500 transition-colors"
                    >
                        {isSidebarOpen ? <ChevronLeft size={20} /> : <Menu size={20} />}
                    </button>
                </div>

                <nav className="flex-1 px-4 space-y-2 mt-4">
                    <NavItem icon={<Mail size={20} />} label="Smart Inbox" active={activeTab === 'inbox'} onClick={() => setActiveTab('inbox')} collapsed={!isSidebarOpen} />
                    <NavItem icon={<Send size={20} />} label="Sent Mail" active={activeTab === 'sent'} onClick={() => setActiveTab('sent')} collapsed={!isSidebarOpen} />
                    <NavItem icon={<AlertTriangle size={20} />} label="Space Optimizer" active={activeTab === 'cleanup'} onClick={() => setActiveTab('cleanup')} collapsed={!isSidebarOpen} />
                    <NavItem icon={<BarChart3 size={20} />} label="Analytics" active={activeTab === 'stats'} onClick={() => setActiveTab('stats')} collapsed={!isSidebarOpen} />
                </nav>

                <div className="p-4 border-t border-slate-100">
                    <button onClick={handleLogout} className="w-full flex items-center justify-center gap-2 bg-slate-100 text-slate-700 py-2.5 rounded-xl hover:bg-slate-200 hover:text-slate-900 transition-colors shadow-sm font-semibold">
                        <LogOut size={18} />
                        {isSidebarOpen && <span>Sign Out</span>}
                    </button>
                </div>
            </aside>

            {/* Main Content */}
            <main className="flex-1 flex flex-col overflow-hidden">
                {/* Header */}
                <header className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-8 shrink-0">
                    <div className="flex items-center gap-6">
                        {!isListOpen && (activeTab === 'inbox' || activeTab === 'sent') && (
                            <button
                                onClick={() => setIsListOpen(true)}
                                className="p-2 bg-blue-50 text-blue-600 rounded-lg hover:bg-blue-100 transition-colors"
                                title="Show email list"
                            >
                                <ChevronRight size={20} />
                            </button>
                        )}
                        <div className="relative w-96">
                            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
                            <input
                                type="text"
                                placeholder="Search emails..."
                                className="w-full pl-10 pr-4 py-2 bg-slate-100 border-transparent rounded-lg focus:bg-white focus:ring-2 focus:ring-blue-500 transition-all outline-none"
                                value={searchQuery}
                                onChange={(e) => setSearchQuery(e.target.value)}
                                onKeyDown={(e) => e.key === 'Enter' && refetch()}
                            />
                        </div>
                    </div>

                    <div className="flex items-center gap-4">
                        {isBatchCategorizing && (
                            <div className="flex items-center gap-2 px-3 py-1 bg-blue-50 text-blue-600 rounded-full text-xs font-bold animate-pulse border border-blue-100">
                                <Loader2 size={12} className="animate-spin" /> Batch Processing...
                            </div>
                        )}
                        <div className="flex items-center gap-2 px-3 py-1.5 bg-slate-100 rounded-lg border border-slate-200">
                            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Limit</span>
                            <select
                                value={limit}
                                onChange={(e) => setLimit(Number(e.target.value))}
                                className="bg-transparent text-sm font-bold text-slate-700 outline-none cursor-pointer"
                            >
                                <option value={20}>20</option>
                                <option value={50}>50</option>
                                <option value={100}>100</option>
                                <option value={200}>200</option>
                            </select>
                        </div>
                        <button onClick={() => refetch()} className="p-2 text-slate-500 hover:bg-slate-100 rounded-full transition-colors relative group">
                            <RefreshCw size={20} className={isLoading ? "animate-spin text-blue-500" : ""} />
                        </button>
                        <div className="w-8 h-8 rounded-full bg-gradient-to-tr from-blue-500 to-indigo-600 shadow-md ring-2 ring-white cursor-pointer relative overflow-hidden flex items-center justify-center text-white font-bold text-xs">
                            Me
                        </div>
                    </div>
                </header>

                {/* View Content */}
                <div className="flex-1 flex overflow-hidden">
                    {(activeTab === 'inbox' || activeTab === 'sent') && (
                        <>
                            {/* Email List */}
                            <div
                                ref={listRef}
                                className={`${isListOpen ? 'w-1/3 xl:w-1/4 opacity-100' : 'w-0 opacity-0 pointer-events-none'} min-w-0 border-r border-slate-200 overflow-y-auto bg-white flex flex-col transition-all duration-300 ease-in-out relative group`}
                            >
                                <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center sticky top-0 z-10 gap-3">
                                    <div className="flex items-center gap-3">
                                        <input
                                            type="checkbox"
                                            className="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                                            checked={emails?.length > 0 && selectedIds.length === emails?.length}
                                            onChange={toggleSelectAll}
                                        />
                                        <h3 className="font-semibold text-slate-800 whitespace-nowrap">
                                            {activeTab === 'sent' ? 'Sent Mail' : 'Inbox Results'}
                                        </h3>
                                    </div>
                                    <div className="flex items-center gap-2">
                                        <span className="text-xs font-bold text-slate-500 bg-slate-200 px-2.5 py-1 rounded-full">{emails?.length || 0}</span>
                                        <button
                                            onClick={() => setIsListOpen(false)}
                                            className="p-1 hover:bg-slate-200 rounded text-slate-400 opacity-0 group-hover:opacity-100 transition-opacity"
                                            title="Collapse list"
                                        >
                                            <ChevronLeft size={14} />
                                        </button>
                                    </div>
                                </div>

                                {selectedIds.length > 0 && (
                                    <div className="mx-4 mt-2 px-4 py-3 bg-gradient-to-r from-blue-600 to-indigo-600 rounded-xl shadow-lg ring-1 ring-white/20 flex items-center justify-between text-white shrink-0">
                                        <div className="flex items-center gap-2">
                                            <div className="w-6 h-6 bg-white/20 rounded-full flex items-center justify-center text-[10px] font-black">{selectedIds.length}</div>
                                            <span className="text-sm font-bold tracking-tight">Selected</span>
                                        </div>
                                        <button
                                            disabled={isBatchCategorizing}
                                            onClick={handleBatchCategorize}
                                            className="flex items-center gap-1.5 px-3 py-1 bg-white/20 hover:bg-white text-white hover:text-blue-700 rounded-lg text-xs font-bold transition-all disabled:opacity-50"
                                        >
                                            {isBatchCategorizing ? <Loader2 size={12} className="animate-spin" /> : <Tag size={12} />}
                                            Categorize
                                        </button>
                                        <button onClick={() => setSelectedIds([])} className="p-1 hover:bg-white/10 rounded-full transition-colors">
                                            <X size={14} />
                                        </button>
                                    </div>
                                )}

                                {isLoading && limit === 20 ? ( // Only show full screen loader for initial load
                                    <div className="flex-1 flex items-center justify-center"><Loader2 className="animate-spin text-blue-500 w-8 h-8" /></div>
                                ) : emails?.length === 0 ? (
                                    <div className="flex-1 flex flex-col items-center justify-center text-slate-400 p-8 text-center">
                                        <Mail size={32} className="mb-3 opacity-20" />
                                        <p>No emails found matching your query.</p>
                                    </div>
                                ) : (
                                    <div className="divide-y divide-slate-100">
                                        {emails?.map((email: any) => (
                                            <div
                                                key={email.id}
                                                onClick={() => handleEmailSelect(email)}
                                                className={`p-5 cursor-pointer hover:bg-slate-50 transition-all flex gap-4 ${selectedEmail?.id === email.id ? 'bg-blue-50/60 border-l-4 border-l-blue-600' : 'border-l-4 border-l-transparent'}`}
                                            >
                                                <div className="pt-1">
                                                    <input
                                                        type="checkbox"
                                                        className="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                                                        checked={selectedIds.includes(email.id)}
                                                        onClick={(e) => e.stopPropagation()}
                                                        onChange={() => toggleSelect(email.id)}
                                                    />
                                                </div>
                                                <div className="flex-1 min-w-0">
                                                    <div className="flex justify-between items-start mb-1.5 gap-2">
                                                        <span className="font-bold text-slate-900 truncate min-w-0 flex-1">
                                                            {activeTab === 'sent' ? (email.to?.split(' <')[0] || 'Recipient') : email.from.split(' <')[0]}
                                                        </span>
                                                        <div className="flex items-center gap-2 shrink-0">
                                                            {activeTab === 'sent' && (
                                                                <span className={`flex items-center gap-1 text-[10px] font-black uppercase px-2 py-0.5 rounded ${email.opened ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-400'}`}>
                                                                    {email.opened ? <><Check size={10} strokeWidth={3} /> Read</> : 'Sent'}
                                                                </span>
                                                            )}
                                                            <span className="text-xs font-medium text-slate-400 whitespace-nowrap">{formatDate(email.date)}</span>
                                                        </div>
                                                    </div>
                                                    <h4 className="text-sm font-semibold text-slate-800 line-clamp-1 mb-1">{email.subject}</h4>
                                                    <p className="text-xs text-slate-500 line-clamp-2 leading-relaxed mb-2">{email.snippet}</p>
                                                    <div className="flex flex-wrap gap-1">
                                                        {email.labels && email.labels.split(', ')
                                                            .filter((l: string) => l.startsWith('MailMaster/'))
                                                            .map((l: string) => (
                                                                <span key={l} className={`text-[10px] px-2 py-0.5 rounded-full font-bold border ${getCategoryColor(l.split('/')[1])}`}>
                                                                    {l.split('/')[1]}
                                                                </span>
                                                            ))}
                                                    </div>
                                                </div>
                                            </div>
                                        ))}
                                        {isLoading && limit > 20 && ( // Show loading more indicator
                                            <div className="p-8 flex justify-center items-center">
                                                <Loader2 className="animate-spin text-blue-500" size={24} />
                                                <span className="ml-3 text-sm font-medium text-slate-500 tracking-wide">Loading more magic...</span>
                                            </div>
                                        )}
                                        {emails && emails.length >= limit && !isLoading && ( // Show load more button only if not loading and more might exist
                                            <div className="p-4 flex justify-center sticky bottom-0 bg-white border-t border-slate-100">
                                                <button
                                                    onClick={() => setLimit(l => l + 20)}
                                                    className="w-full text-sm font-bold text-slate-700 hover:text-blue-700 bg-slate-50 hover:bg-blue-50 px-5 py-3 rounded-xl transition-all border border-slate-200 hover:border-blue-200 shadow-sm"
                                                >
                                                    Load Older Emails
                                                </button>
                                            </div>
                                        )}
                                    </div>
                                )}
                            </div>

                            {/* Email Reading Detail */}
                            <div className="flex-1 bg-[#FAFBFF] p-8 overflow-y-auto min-w-0">
                                {selectedEmail ? (
                                    <div className="max-w-4xl mx-auto w-full bg-white border border-slate-200 rounded-2xl shadow-sm p-6 md:p-8 min-h-[80vh] flex flex-col">
                                        <div className="flex flex-col xl:flex-row justify-between items-start mb-8 pb-6 border-b border-slate-100 gap-6 w-full">
                                            <div className="min-w-0 flex-1 w-full relative">
                                                <h2 className="text-2xl md:text-3xl font-bold text-slate-900 mb-3 leading-tight break-words" style={{ wordBreak: 'break-word' }}>{selectedEmail.subject}</h2>
                                                <div className="flex items-center gap-3">
                                                    <div className="flex-shrink-0">
                                                        <div className="w-10 h-10 rounded-full bg-slate-100 flex items-center justify-center text-slate-600 font-bold">
                                                            {(activeTab === 'sent' ? (selectedEmail.to?.[0] || '?') : (selectedEmail.from?.[0] || '?')).toUpperCase()}
                                                        </div>
                                                    </div>
                                                    <div className="min-w-0 flex-1">
                                                        <span className="font-semibold text-slate-800 block truncate">
                                                            {activeTab === 'sent' ? `Sent To: ${selectedEmail.to}` : `From: ${selectedEmail.from}`}
                                                        </span>
                                                        <div className="flex items-center gap-2 mt-1">
                                                            <span className="text-xs text-slate-500">{formatDate(selectedEmail.date)}</span>
                                                            {activeTab === 'sent' && (
                                                                <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${selectedEmail.opened ? 'bg-emerald-50 text-emerald-600 border-emerald-100' : 'bg-slate-50 text-slate-400 border-slate-200'}`}>
                                                                    {selectedEmail.opened ? '✓ Read' : '• Pending'}
                                                                </span>
                                                            )}
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="flex gap-3 shrink-0 flex-wrap">
                                                <button
                                                    onClick={() => handleSummarize(selectedEmail.id)}
                                                    disabled={isSummarizing}
                                                    className={`flex items-center gap-2 px-4 py-2 rounded-xl transition-all font-semibold shadow-sm text-sm ${isSummarizing ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : 'bg-gradient-to-r from-blue-50 to-indigo-50 text-indigo-700 border border-indigo-100 hover:from-blue-100 hover:to-indigo-100 cursor-pointer'}`}
                                                >
                                                    {isSummarizing ? <Loader2 size={16} className="animate-spin" /> : <Sparkles size={16} className="text-indigo-500" />}
                                                    {isSummarizing ? "Working..." : "AI Summary"}
                                                </button>
                                                <button
                                                    onClick={() => handleCategorize(selectedEmail.id, selectedEmail)}
                                                    disabled={isCategorizing}
                                                    className={`flex items-center gap-2 px-4 py-2 rounded-xl transition-all font-semibold text-sm shadow-sm ${isCategorizing ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : 'bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 cursor-pointer'}`}
                                                >
                                                    {isCategorizing ? <Loader2 size={16} className="animate-spin" /> : <Tag size={16} className="text-slate-400" />}
                                                    {isCategorizing ? "Thinking..." : "Categorize"}
                                                </button>
                                            </div>
                                        </div>

                                        <div className="flex flex-wrap gap-2 mb-6">
                                            {selectedEmail.labels && selectedEmail.labels.split(', ')
                                                .filter((l: string) => !['INBOX', 'UNREAD', 'IMPORTANT', 'CATEGORY_PROMOTIONS', 'CATEGORY_UPDATES', 'CATEGORY_SOCIAL', 'CATEGORY_FORUMS', 'CHAT'].includes(l))
                                                .map((label: string) => (
                                                    <div
                                                        key={label}
                                                        className={`inline-flex items-center gap-2 pl-3 pr-1.5 py-1.5 rounded-full text-xs font-bold uppercase tracking-wider border shadow-sm group ${getCategoryColor(label.includes('/') ? label.split('/')[1] : label)}`}
                                                    >
                                                        <span className={`w-2 h-2 rounded-full ${label.includes('Priority') ? 'bg-blue-500' : 'bg-slate-400'}`}></span>
                                                        {label.startsWith('MailMaster/') ? label.replace('MailMaster/', '') : label}
                                                        <button
                                                            onClick={() => handleRemoveLabel(selectedEmail.id, label)}
                                                            className="p-0.5 hover:bg-black/5 rounded-full transition-colors ml-1"
                                                            title="Remove label and return to inbox"
                                                        >
                                                            <X size={14} />
                                                        </button>
                                                    </div>
                                                ))}

                                            {selectedEmail.category && !selectedEmail.labels?.includes(`MailMaster/${selectedEmail.category}`) && (
                                                <div className={`inline-flex items-center gap-2 px-3 py-1.5 rounded-full text-xs font-bold uppercase border animate-in fade-in zoom-in duration-300 ${getCategoryColor(selectedEmail.category)}`}>
                                                    <Check size={14} /> {selectedEmail.category}
                                                </div>
                                            )}
                                        </div>

                                        {selectedEmail.summary && (
                                            <div className="bg-slate-50 p-6 rounded-2xl border border-slate-200 mb-8 relative overflow-hidden">
                                                <div className="absolute top-0 left-0 w-1.5 h-full bg-gradient-to-b from-blue-400 to-indigo-600"></div>
                                                <h5 className="text-sm font-bold text-slate-800 mb-3 flex items-center gap-2">
                                                    <Sparkles size={16} className="text-blue-500" /> Executive Summary
                                                </h5>
                                                <p className="text-slate-700 leading-relaxed font-medium mb-4">"{selectedEmail.summary}"</p>

                                                {selectedEmail.reasoning && (
                                                    <div className="flex items-start gap-2 text-xs text-slate-400 bg-white/50 p-2 rounded-lg border border-slate-100">
                                                        <Info size={14} className="shrink-0 mt-0.5" />
                                                        <span className="italic">{selectedEmail.reasoning.split('Reasoning: ')[1] || selectedEmail.reasoning}</span>
                                                    </div>
                                                )}
                                            </div>
                                        )}

                                        {selectedEmail.attachments && selectedEmail.attachments.length > 0 && (
                                            <div className="mb-8 p-4 bg-slate-50 border border-slate-200 rounded-2xl">
                                                <h5 className="text-sm font-bold text-slate-800 mb-3 flex items-center gap-2">
                                                    <Paperclip size={16} className="text-slate-400" /> Attachments ({selectedEmail.attachments.length})
                                                </h5>
                                                <div className="flex flex-wrap gap-2">
                                                    {selectedEmail.attachments.map((att: any) => (
                                                        <a
                                                            key={att.id}
                                                            href={`${API_BASE}/emails/${selectedEmail.id}/attachments/${att.id}?filename=${encodeURIComponent(att.filename)}`}
                                                            target="_blank"
                                                            rel="noopener noreferrer"
                                                            className="flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-semibold text-slate-700 hover:bg-blue-50 hover:border-blue-200 hover:text-blue-700 transition-all"
                                                        >
                                                            <Download size={14} className="opacity-60" />
                                                            {att.filename}
                                                            <span className="text-[10px] text-slate-400">({(att.size / 1024).toFixed(1)} KB)</span>
                                                        </a>
                                                    ))}
                                                </div>
                                            </div>
                                        )}

                                        <div className="flex justify-between items-center mb-6">
                                            <h5 className="text-sm font-bold text-slate-400 uppercase tracking-widest">Email Content</h5>
                                            <button
                                                onClick={() => setExpandedEmails(prev => ({ ...prev, [selectedEmail.id]: !prev[selectedEmail.id] }))}
                                                className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-lg text-xs font-bold transition-all"
                                            >
                                                {expandedEmails[selectedEmail.id] ? (
                                                    <>Less <ChevronUp size={14} /></>
                                                ) : (
                                                    <>More <ChevronDown size={14} /></>
                                                )}
                                            </button>
                                        </div>

                                        <div className={`text-slate-800 leading-relaxed overflow-x-auto ${expandedEmails[selectedEmail.id] ? '' : 'max-h-[300px] relative'}`}>
                                            {/* We use sanitized div for HTML content */}
                                            {selectedEmail.body && selectedEmail.body.includes('<') ? (
                                                <div className="email-body-container">
                                                    <div
                                                        className="prose prose-slate max-w-none"
                                                        dangerouslySetInnerHTML={{ __html: selectedEmail.body }}
                                                    />
                                                </div>
                                            ) : (
                                                <div className="whitespace-pre-wrap font-sans text-[15px]">
                                                    {selectedEmail.body || selectedEmail.snippet}
                                                </div>
                                            )}

                                            {!expandedEmails[selectedEmail.id] && (
                                                <div className="absolute bottom-0 left-0 w-full h-24 bg-gradient-to-t from-white to-transparent pointer-events-none"></div>
                                            )}
                                        </div>
                                    </div>
                                ) : (
                                    <div className="h-full flex flex-col items-center justify-center text-slate-400 gap-6">
                                        <div className="w-24 h-24 bg-slate-100 rounded-full flex items-center justify-center">
                                            <Mail size={40} className="text-slate-300" />
                                        </div>
                                        <p className="text-lg font-medium text-slate-500">Select an email to dive in</p>
                                    </div>
                                )}
                            </div>
                        </>
                    )}

                    {activeTab === 'cleanup' && (
                        <div className="flex-1 bg-[#FAFBFF] p-8 overflow-y-auto">
                            <div className="max-w-4xl mx-auto">
                                <div className="mb-10 text-center">
                                    <div className="w-16 h-16 bg-red-50 text-red-500 flex items-center justify-center rounded-full mx-auto mb-4">
                                        <Trash2 size={32} />
                                    </div>
                                    <h2 className="text-3xl font-bold text-slate-900 mb-3">Space Optimizer</h2>
                                    <p className="text-slate-500 max-w-lg mx-auto">We've identified the largest files hiding in your inbox. Review and remove them to instantly free up your Google Drive storage.</p>
                                </div>

                                {isLoadingCleanup ? (
                                    <div className="flex justify-center p-12"><Loader2 className="animate-spin text-red-500 w-10 h-10" /></div>
                                ) : (
                                    <div className="grid grid-cols-1 gap-4">
                                        {cleanupData?.large_emails?.map((email: any) => (
                                            <div key={email.id} className="bg-white p-5 rounded-2xl border border-slate-200 flex items-center justify-between hover:shadow-lg hover:border-red-100 transition-all group">
                                                <div className="flex gap-5 items-center">
                                                    <div className="w-12 h-12 bg-slate-50 rounded-xl flex items-center justify-center text-slate-400 group-hover:bg-red-50 group-hover:text-red-500 transition-colors">
                                                        <Mail size={24} />
                                                    </div>
                                                    <div>
                                                        <h4 className="font-bold text-slate-900 text-lg mb-1">{email.subject}</h4>
                                                        <p className="text-sm text-slate-500 font-medium">From {email.sender.split('<')[0]} <span className="mx-2 text-slate-300">•</span> {formatDate(email.date)}</p>
                                                    </div>
                                                </div>
                                                <div className="flex items-center gap-6">
                                                    <div className="text-right">
                                                        <span className="block text-xl font-black text-slate-800">{email.size_mb.toFixed(1)} <span className="text-sm font-semibold text-slate-500">MB</span></span>
                                                    </div>
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>
                        </div>
                    )}
                </div>
            </main>

            {/* Compose FAB */}
            <button
                onClick={() => setIsComposeOpen(true)}
                className="fixed bottom-8 right-8 w-14 h-14 bg-gradient-to-r from-blue-600 to-indigo-600 rounded-full shadow-xl shadow-blue-500/30 flex items-center justify-center text-white hover:scale-105 active:scale-95 transition-all z-50 group hover:shadow-2xl hover:shadow-blue-500/40 border border-white/10"
                title="Compose Email"
            >
                <Edit2 size={24} className="group-hover:rotate-12 transition-transform duration-300" />
            </button>

            {/* Compose Modal */}
            {isComposeOpen && (
                <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 sm:p-6 pb-20">
                    <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm" onClick={() => !isSending && setIsComposeOpen(false)}></div>
                    <div className="relative w-full max-w-2xl bg-white rounded-2xl shadow-2xl border border-slate-100 flex flex-col overflow-hidden animate-in fade-in slide-in-from-bottom-8 duration-300">
                        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100 bg-slate-50/50">
                            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                                <Edit2 size={18} className="text-blue-600" /> New Message
                            </h3>
                            <button onClick={() => !isSending && setIsComposeOpen(false)} className="p-1 hover:bg-slate-200 rounded text-slate-500 transition-colors">
                                <X size={20} />
                            </button>
                        </div>
                        <div className="p-6 space-y-4 flex-1">
                            <div>
                                <input
                                    type="email"
                                    placeholder="To"
                                    className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all outline-none font-medium"
                                    value={composeForm.to}
                                    onChange={(e) => setComposeForm(prev => ({ ...prev, to: e.target.value }))}
                                />
                            </div>
                            <div>
                                <input
                                    type="text"
                                    placeholder="Subject"
                                    className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all outline-none font-bold text-slate-800"
                                    value={composeForm.subject}
                                    onChange={(e) => setComposeForm(prev => ({ ...prev, subject: e.target.value }))}
                                />
                            </div>
                            <div className="h-64">
                                <textarea
                                    placeholder="Write your email here..."
                                    className="w-full h-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all outline-none resize-none prose prose-slate"
                                    value={composeForm.body}
                                    onChange={(e) => setComposeForm(prev => ({ ...prev, body: e.target.value }))}
                                ></textarea>
                            </div>
                        </div>
                        <div className="px-6 py-4 border-t border-slate-100 bg-slate-50 flex justify-end gap-3">
                            <button
                                onClick={() => setIsComposeOpen(false)}
                                disabled={isSending}
                                className="px-5 py-2 rounded-xl text-slate-600 font-bold hover:bg-slate-200 transition-colors"
                            >
                                Cancel
                            </button>
                            <button
                                onClick={handleSendEmail}
                                disabled={isSending || !composeForm.to || !composeForm.subject}
                                className="flex items-center gap-2 px-6 py-2 bg-gradient-to-r from-blue-600 to-indigo-600 text-white rounded-xl font-bold hover:shadow-lg transition-all disabled:opacity-50 disabled:cursor-not-allowed"
                            >
                                {isSending ? <Loader2 size={18} className="animate-spin" /> : <Send size={18} />}
                                {isSending ? 'Sending...' : 'Send'}
                            </button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}

function NavItem({ icon, label, active, onClick, collapsed }: any) {
    return (
        <button
            onClick={onClick}
            className={`w-full flex items-center ${collapsed ? 'justify-center' : 'gap-3 px-4'} py-3.5 rounded-xl transition-all font-semibold ${active ? 'bg-blue-600 text-white shadow-md shadow-blue-600/20' : 'text-slate-500 hover:bg-slate-100 hover:text-slate-900'}`}
        >
            {icon}
            {!collapsed && <span>{label}</span>}
        </button>
    );
}
