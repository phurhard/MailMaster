import { useState } from 'react';
import { useQuery } from '@tanstack/react-query';
import { useLocation } from 'wouter';
import {
    Mail, Trash2, Sparkles, Settings as SettingsIcon,
    Search, AlertTriangle, BarChart3, RefreshCw, LogOut, Loader2
} from 'lucide-react';

const API_BASE = 'http://localhost:8000';

export default function Dashboard() {
    const [, setLocation] = useLocation();
    const [activeTab, setActiveTab] = useState('inbox');
    const [searchQuery, setSearchQuery] = useState('');
    const [limit, setLimit] = useState(20);
    const [selectedEmail, setSelectedEmail] = useState<any>(null);
    const [isSidebarOpen] = useState(true);

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
        queryKey: ['emails', searchQuery, limit],
        queryFn: async () => {
            const endpoint = searchQuery
                ? `${API_BASE}/emails/search/${searchQuery}?limit=${limit}`
                : `${API_BASE}/emails/search/in:inbox?limit=${limit}`;
            const res = await authorizedFetch(endpoint);
            return res.json();
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
        const res = await authorizedFetch(`${API_BASE}/emails/${id}/summarize`, { method: 'POST' });
        const data = await res.json();
        setSelectedEmail((prev: any) => ({ ...prev, summary: data.summary }));
    };

    const handleCategorize = async (id: string) => {
        const res = await authorizedFetch(`${API_BASE}/emails/${id}/categorize`, { method: 'POST' });
        const data = await res.json();
        setSelectedEmail((prev: any) => ({ ...prev, category: data.category }));
    };

    return (
        <div className="flex h-screen bg-slate-50 font-sans text-slate-900">
            {/* Sidebar */}
            <aside className={`${isSidebarOpen ? 'w-64' : 'w-20'} bg-white border-r border-slate-200 transition-all duration-300 flex flex-col`}>
                <div className="p-6 flex items-center gap-3">
                    <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white">
                        <Mail size={20} />
                    </div>
                    {isSidebarOpen && <span className="font-bold text-xl tracking-tight">MailMaster</span>}
                </div>

                <nav className="flex-1 px-4 space-y-2 mt-4">
                    <NavItem icon={<Mail size={20} />} label="Smart Inbox" active={activeTab === 'inbox'} onClick={() => setActiveTab('inbox')} collapsed={!isSidebarOpen} />
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

                    <div className="flex items-center gap-4">
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
                    {activeTab === 'inbox' && (
                        <>
                            {/* Email List */}
                            <div className="w-1/3 xl:w-1/4 min-w-[320px] border-r border-slate-200 overflow-y-auto bg-white flex flex-col">
                                <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center sticky top-0 z-10">
                                    <h3 className="font-semibold text-slate-800">Inbox Results</h3>
                                    <span className="text-xs font-bold text-slate-500 bg-slate-200 px-2.5 py-1 rounded-full">{emails?.length || 0}</span>
                                </div>
                                {isLoading ? (
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
                                                onClick={() => setSelectedEmail(email)}
                                                className={`p-5 cursor-pointer hover:bg-slate-50 transition-all ${selectedEmail?.id === email.id ? 'bg-blue-50/60 border-l-4 border-l-blue-600' : 'border-l-4 border-l-transparent'}`}
                                            >
                                                <div className="flex justify-between items-center mb-1.5">
                                                    <span className="font-bold text-slate-900 truncate pr-2">{email.from.split(' <')[0]}</span>
                                                    <span className="text-xs font-medium text-slate-400 whitespace-nowrap">{email.date}</span>
                                                </div>
                                                <h4 className="text-sm font-semibold text-slate-800 line-clamp-1 mb-1">{email.subject}</h4>
                                                <p className="text-xs text-slate-500 line-clamp-2 leading-relaxed">{email.snippet}</p>
                                            </div>
                                        ))}
                                        {emails && emails.length >= limit && (
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
                            <div className="flex-1 bg-[#FAFBFF] p-8 overflow-y-auto">
                                {selectedEmail ? (
                                    <div className="max-w-3xl mx-auto bg-white border border-slate-200 rounded-2xl shadow-sm p-8 min-h-[80vh]">
                                        <div className="flex justify-between items-start mb-8 pb-6 border-b border-slate-100">
                                            <div>
                                                <h2 className="text-3xl font-bold text-slate-900 mb-3 leading-tight">{selectedEmail.subject}</h2>
                                                <div className="flex items-center gap-3">
                                                    <div className="w-10 h-10 rounded-full bg-slate-100 flex items-center justify-center text-slate-600 font-bold">
                                                        {selectedEmail.from.charAt(0).toUpperCase()}
                                                    </div>
                                                    <div>
                                                        <span className="font-semibold text-slate-800 block">{selectedEmail.from.split('<')[0]}</span>
                                                        <span className="text-xs text-slate-500">{selectedEmail.date}</span>
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="flex gap-3">
                                                <button onClick={() => handleSummarize(selectedEmail.id)} className="flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-blue-50 to-indigo-50 text-indigo-700 border border-indigo-100 rounded-xl hover:from-blue-100 hover:to-indigo-100 transition-all font-semibold shadow-sm text-sm">
                                                    <Sparkles size={16} className="text-indigo-500" /> AI Summary
                                                </button>
                                                <button onClick={() => handleCategorize(selectedEmail.id)} className="flex items-center gap-2 px-4 py-2 bg-white text-slate-700 border border-slate-200 rounded-xl hover:bg-slate-50 transition-all font-semibold text-sm shadow-sm">
                                                    <SettingsIcon size={16} className="text-slate-400" /> Categorize
                                                </button>
                                            </div>
                                        </div>

                                        {selectedEmail.category && (
                                            <div className="mb-6 inline-flex items-center gap-2 px-4 py-1.5 bg-indigo-50 text-indigo-700 rounded-full text-xs font-bold uppercase tracking-wider border border-indigo-100">
                                                <span className="w-2 h-2 rounded-full bg-indigo-500"></span>
                                                {selectedEmail.category}
                                            </div>
                                        )}

                                        {selectedEmail.summary && (
                                            <div className="bg-slate-50 p-6 rounded-2xl border border-slate-200 mb-8 relative overflow-hidden">
                                                <div className="absolute top-0 left-0 w-1.5 h-full bg-gradient-to-b from-blue-400 to-indigo-600"></div>
                                                <h5 className="text-sm font-bold text-slate-800 mb-3 flex items-center gap-2">
                                                    <Sparkles size={16} className="text-blue-500" /> Executive Summary
                                                </h5>
                                                <p className="text-slate-700 leading-relaxed font-medium">"{selectedEmail.summary}"</p>
                                            </div>
                                        )}

                                        <div className="text-slate-800 leading-relaxed whitespace-pre-wrap text-[15px]">
                                            {selectedEmail.snippet}
                                            <br /><br />
                                            <span className="text-slate-400 italic text-sm">(Preview shown. HTML rendering to be implemented.)</span>
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
                                                        <p className="text-sm text-slate-500 font-medium">From {email.sender.split('<')[0]} <span className="mx-2 text-slate-300">•</span> {email.date}</p>
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
