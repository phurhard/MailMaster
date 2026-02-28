import { useEffect } from 'react';
import { useLocation } from 'wouter';
import { Loader2 } from 'lucide-react';

export default function OAuthCallback() {
    const [, setLocation] = useLocation();

    useEffect(() => {
        // Extract the token exactly matching what we sent from FastAPI
        // Example: ?token=ey...
        const urlParams = new URLSearchParams(window.location.search);
        const token = urlParams.get('token');

        if (token) {
            localStorage.setItem('auth_token', token);

            // Delay briefly so the user sees a smooth transition, then redirect
            setTimeout(() => {
                setLocation('/');
            }, 800);
        } else {
            // If no token was found, redirect back to login
            setLocation('/login');
        }
    }, [setLocation]);

    return (
        <div className="min-h-screen bg-slate-900 flex flex-col items-center justify-center p-4">
            <div className="relative">
                {/* Glow effect */}
                <div className="absolute inset-0 bg-blue-500 blur-[80px] opacity-20 rounded-full"></div>

                <div className="relative bg-white/5 backdrop-blur-xl border border-white/10 p-10 rounded-3xl shadow-2xl flex flex-col items-center gap-6">
                    <div className="bg-blue-600/20 p-4 rounded-2xl">
                        <Loader2 className="w-10 h-10 text-blue-400 animate-spin" />
                    </div>
                    <div className="text-center">
                        <h2 className="text-2xl font-bold text-white mb-2">Authenticating</h2>
                        <p className="text-slate-400 font-medium">Securing your connection to MailMaster...</p>
                    </div>
                </div>
            </div>
        </div>
    );
}
