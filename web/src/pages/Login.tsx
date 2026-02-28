import { Mail, ShieldCheck, Zap, Lock, ArrowRight } from 'lucide-react';

const API_BASE = 'http://localhost:8000';

export default function Login() {
    const handleLogin = () => {
        window.location.href = `${API_BASE}/auth/login`;
    };

    return (
        <div className="relative min-h-screen overflow-hidden bg-[#0A0F1C] text-white flex items-center justify-center p-4 font-sans border-box">
            {/* Dynamic Background Elements */}
            <div className="absolute top-[-10%] left-[-10%] w-[500px] h-[500px] bg-indigo-600/30 rounded-full blur-[120px] mix-blend-screen opacity-60 animate-blob"></div>
            <div className="absolute bottom-[-10%] right-[-10%] w-[600px] h-[600px] bg-blue-600/20 rounded-full blur-[150px] mix-blend-screen opacity-50 animate-blob animation-delay-2000"></div>

            <div className="relative w-full max-w-[1100px] grid grid-cols-1 lg:grid-cols-2 gap-12 items-center z-10">

                {/* Left Side: Copy & Branding */}
                <div className="flex flex-col gap-8 px-4 lg:px-0">
                    <div className="inline-flex items-center gap-3 bg-white/5 border border-white/10 backdrop-blur-md px-4 py-2 rounded-full w-fit">
                        <Mail className="w-5 h-5 text-blue-400" />
                        <span className="text-sm font-semibold tracking-wide text-blue-100">MailMaster AI v1.0</span>
                    </div>

                    <h1 className="text-5xl lg:text-7xl font-bold tracking-tight leading-[1.1]">
                        Your inbox, <br />
                        <span className="text-transparent bg-clip-text bg-gradient-to-r from-blue-400 via-indigo-400 to-purple-400">
                            reimagined.
                        </span>
                    </h1>

                    <p className="text-lg text-slate-400 leading-relaxed max-w-xl font-medium">
                        Connect your Gmail account to unlock AI-powered summaries, instant categorization, and smart cleanup suggestions. Experience the future of email management.
                    </p>

                    <div className="flex flex-col sm:flex-row gap-4 mt-4">
                        <div className="flex items-center gap-3">
                            <div className="bg-emerald-500/10 p-2 rounded-lg">
                                <ShieldCheck className="w-5 h-5 text-emerald-400" />
                            </div>
                            <span className="text-sm font-medium text-slate-300">OAuth 2.0 Secured</span>
                        </div>
                        <div className="flex items-center gap-3">
                            <div className="bg-purple-500/10 p-2 rounded-lg">
                                <Zap className="w-5 h-5 text-purple-400" />
                            </div>
                            <span className="text-sm font-medium text-slate-300">AI Powered Analysis</span>
                        </div>
                    </div>
                </div>

                {/* Right Side: Auth Card */}
                <div className="relative">
                    {/* Card glow */}
                    <div className="absolute -inset-1 bg-gradient-to-r from-blue-500 to-purple-600 rounded-[2.5rem] blur-xl opacity-20 group-hover:opacity-40 transition duration-1000"></div>

                    <div className="relative bg-white/[0.03] backdrop-blur-2xl border border-white/10 rounded-[2rem] p-10 lg:p-12 shadow-2xl flex flex-col items-center text-center">

                        <div className="w-20 h-20 bg-gradient-to-br from-blue-500 to-indigo-600 rounded-2xl flex items-center justify-center mb-8 shadow-inner shadow-white/20">
                            <Lock className="w-10 h-10 text-white" />
                        </div>

                        <h3 className="text-3xl font-bold text-white mb-3">Welcome Back</h3>
                        <p className="text-slate-400 mb-10 font-medium">Log in to safely sync and analyze your latest emails.</p>

                        <button
                            onClick={handleLogin}
                            className="group relative w-full flex items-center justify-center gap-4 bg-white text-slate-900 px-8 py-4 rounded-xl font-bold text-lg hover:bg-slate-100 hover:scale-[1.02] transition-all duration-300"
                        >
                            {/* Google Logo SVG */}
                            <svg className="w-6 h-6" viewBox="0 0 24 24">
                                <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.09-1.92 3.28-4.74 3.28-8.09z" />
                                <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" />
                                <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z" />
                                <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" />
                            </svg>
                            <span>Continue with Google</span>
                            <ArrowRight className="w-5 h-5 absolute right-6 opacity-0 -translate-x-4 group-hover:opacity-100 group-hover:translate-x-0 transition-all duration-300" />
                        </button>

                        <p className="mt-8 text-xs text-slate-500 max-w-xs leading-relaxed">
                            By continuing, you agree to our Terms of Service and Privacy Policy. We only request permissions necessary to analyze your inbox.
                        </p>
                    </div>
                </div>

            </div>
        </div>
    );
}
