import { useState, useEffect } from 'react'
import {
  Mail,
  Trash2,
  Sparkles,
  ShieldCheck,
  Search,
  Menu,
  X,
  ChevronRight,
  BarChart3,
  Settings,
  AlertTriangle,
  RefreshCw
} from 'lucide-react'
import { useQuery } from '@tanstack/react-query'

const API_BASE = 'http://localhost:8000'

function App() {
  const [activeTab, setActiveTab] = useState('inbox')
  const [searchQuery, setSearchQuery] = useState('')
  const [selectedEmail, setSelectedEmail] = useState<any>(null)
  const [isSidebarOpen, setIsSidebarOpen] = useState(true)

  // Fetch emails
  const { data: emails, isLoading, refetch } = useQuery({
    queryKey: ['emails', searchQuery],
    queryFn: async () => {
      const endpoint = searchQuery
        ? `${API_BASE}/emails/search/${searchQuery}`
        : `${API_BASE}/emails/search/in:inbox`
      const res = await fetch(endpoint)
      return res.json()
    }
  })

  // Fetch cleanup suggestions
  const { data: cleanupData } = useQuery({
    queryKey: ['cleanup'],
    queryFn: async () => {
      const res = await fetch(`${API_BASE}/emails/cleanup-suggestions`)
      return res.json()
    }
  })

  const handleLogin = () => {
    window.location.href = `${API_BASE}/auth/login`
  }

  const handleSummarize = async (id: string) => {
    const res = await fetch(`${API_BASE}/emails/${id}/summarize`, { method: 'POST' })
    const data = await res.json()
    setSelectedEmail((prev: any) => ({ ...prev, summary: data.summary }))
  }

  const handleCategorize = async (id: string) => {
    const res = await fetch(`${API_BASE}/emails/${id}/categorize`, { method: 'POST' })
    const data = await res.json()
    setSelectedEmail((prev: any) => ({ ...prev, category: data.category }))
  }

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
          <NavItem
            icon={<Mail size={20} />}
            label="Smart Inbox"
            active={activeTab === 'inbox'}
            onClick={() => setActiveTab('inbox')}
            collapsed={!isSidebarOpen}
          />
          <NavItem
            icon={<AlertTriangle size={20} />}
            label="Space Optimizer"
            active={activeTab === 'cleanup'}
            onClick={() => setActiveTab('cleanup')}
            collapsed={!isSidebarOpen}
          />
          <NavItem
            icon={<BarChart3 size={20} />}
            label="Analytics"
            active={activeTab === 'stats'}
            onClick={() => setActiveTab('stats')}
            collapsed={!isSidebarOpen}
          />
        </nav>

        <div className="p-4 border-t border-slate-100">
          <button
            onClick={handleLogin}
            className="w-full flex items-center justify-center gap-2 bg-slate-900 text-white py-2.5 rounded-xl hover:bg-slate-800 transition-colors shadow-sm"
          >
            <ShieldCheck size={18} />
            {isSidebarOpen && <span>Connect Gmail</span>}
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
            <button className="p-2 text-slate-500 hover:bg-slate-100 rounded-full">
              <RefreshCw size={20} />
            </button>
            <div className="w-8 h-8 rounded-full bg-gradient-to-tr from-blue-500 to-indigo-600 shadow-md"></div>
          </div>
        </header>

        {/* View Content */}
        <div className="flex-1 flex overflow-hidden">
          {activeTab === 'inbox' && (
            <>
              {/* Email List */}
              <div className="w-1/2 border-r border-slate-200 overflow-y-auto bg-white">
                {isLoading ? (
                  <div className="p-8 flex justify-center"><div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div></div>
                ) : (
                  emails?.map((email: any) => (
                    <div
                      key={email.id}
                      onClick={() => setSelectedEmail(email)}
                      className={`p-4 border-b border-slate-100 cursor-pointer hover:bg-slate-50 transition-colors ${selectedEmail?.id === email.id ? 'bg-blue-50/50 border-l-4 border-l-blue-600' : ''}`}
                    >
                      <div className="flex justify-between items-start mb-1">
                        <span className="font-semibold text-slate-800 truncate">{email.from}</span>
                        <span className="text-xs text-slate-400">{email.date}</span>
                      </div>
                      <h4 className="text-sm font-medium text-slate-700 truncate mb-1">{email.subject}</h4>
                      <p className="text-xs text-slate-500 line-clamp-1">{email.snippet}</p>
                    </div>
                  ))
                )}
              </div>

              {/* Email Reading Detail */}
              <div className="flex-1 bg-white p-8 overflow-y-auto">
                {selectedEmail ? (
                  <div className="max-w-2xl mx-auto">
                    <div className="flex justify-between items-start mb-8">
                      <div>
                        <h2 className="text-2xl font-bold text-slate-900 mb-2">{selectedEmail.subject}</h2>
                        <span className="text-slate-500">From: {selectedEmail.from}</span>
                      </div>
                      <div className="flex gap-2">
                        <button
                          onClick={() => handleSummarize(selectedEmail.id)}
                          className="flex items-center gap-2 px-4 py-2 bg-blue-50 text-blue-600 rounded-lg hover:bg-blue-100 transition-colors"
                        >
                          <Sparkles size={18} />
                          <span>AI Summary</span>
                        </button>
                        <button
                          onClick={() => handleCategorize(selectedEmail.id)}
                          className="flex items-center gap-2 px-4 py-2 bg-purple-50 text-purple-600 rounded-lg hover:bg-purple-100 transition-colors"
                        >
                          <Settings size={18} />
                          <span>Categorize</span>
                        </button>
                      </div>
                    </div>

                    {selectedEmail.category && (
                      <div className="mb-4 inline-block px-3 py-1 bg-indigo-100 text-indigo-700 rounded-full text-xs font-bold uppercase tracking-wider">
                        Category: {selectedEmail.category}
                      </div>
                    )}

                    {selectedEmail.summary && (
                      <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 mb-6 relative overflow-hidden">
                        <div className="absolute top-0 left-0 w-1 h-full bg-blue-500"></div>
                        <h5 className="text-xs font-bold text-blue-600 uppercase mb-2 flex items-center gap-2">
                          <Sparkles size={12} /> AI Summary
                        </h5>
                        <p className="text-slate-700 italic">"{selectedEmail.summary}"</p>
                      </div>
                    )}

                    <div className="text-slate-800 leading-relaxed whitespace-pre-wrap">
                      {selectedEmail.snippet}...
                    </div>
                  </div>
                ) : (
                  <div className="h-full flex flex-col items-center justify-center text-slate-400 gap-4">
                    <Mail size={48} className="opacity-20" />
                    <p>Select an email to read and analyze</p>
                  </div>
                )}
              </div>
            </>
          )}

          {activeTab === 'cleanup' && (
            <div className="flex-1 p-8 overflow-y-auto">
              <div className="max-w-4xl mx-auto">
                <div className="mb-8">
                  <h2 className="text-3xl font-bold text-slate-900 mb-2">Space Optimizer</h2>
                  <p className="text-slate-500">Identity and remove large emails to save space on your Google account.</p>
                </div>

                <div className="grid grid-cols-1 gap-6">
                  {cleanupData?.large_emails?.map((email: any) => (
                    <div key={email.id} className="bg-white p-6 rounded-2xl border border-slate-200 flex items-center justify-between hover:shadow-md transition-shadow">
                      <div className="flex gap-4 items-center">
                        <div className="w-12 h-12 bg-orange-50 rounded-xl flex items-center justify-center text-orange-600">
                          <Trash2 size={24} />
                        </div>
                        <div>
                          <h4 className="font-semibold text-slate-900">{email.subject}</h4>
                          <p className="text-sm text-slate-500">Sent by {email.sender} • {new Date(email.date).toLocaleDateString()}</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-6">
                        <span className="text-lg font-bold text-red-600">{email.size_mb.toFixed(1)} MB</span>
                        <button className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-all">
                          <Trash2 size={20} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  )
}

function NavItem({ icon, label, active, onClick, collapsed }: any) {
  return (
    <button
      onClick={onClick}
      className={`w-full flex items-center ${collapsed ? 'justify-center' : 'gap-3 px-4'} py-3 rounded-xl transition-all ${active ? 'bg-blue-50 text-blue-600 shadow-sm shadow-blue-100' : 'text-slate-500 hover:bg-slate-50'}`}
    >
      {icon}
      {!collapsed && <span className="font-medium text-sm">{label}</span>}
    </button>
  )
}

export default App
