import React, { useState, useEffect } from 'react';
import { ShieldCheck, Lock, Eye, EyeOff, KeyRound, Mail } from 'lucide-react';
import { setupPassword, verifyPassword } from '../utils/crypto';

interface LockScreenProps {
  onUnlock: (key: CryptoKey) => void;
}

export const LockScreen: React.FC<LockScreenProps> = ({ onUnlock }) => {
  const [isFirstTime, setIsFirstTime] = useState(false);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isWiggling, setIsWiggling] = useState(false);
  const [isLoading, setIsLoading] = useState(false);

  useEffect(() => {
    const salt = localStorage.getItem('muhasebe_salt');
    const hash = localStorage.getItem('muhasebe_hash');
    if (!salt || !hash) {
      setIsFirstTime(true);
    }
  }, []);

  const triggerWiggle = () => {
    setIsWiggling(true);
    setTimeout(() => setIsWiggling(false), 500);
  };

  const handleAction = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setIsLoading(true);

    try {
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(email)) {
        setError('Lütfen geçerli bir e-posta adresi girin!');
        triggerWiggle();
        setIsLoading(false);
        return;
      }

      if (isFirstTime) {
        // Setup Password
        if (password.length < 6) {
          setError('Master şifre en az 6 karakter olmalıdır!');
          triggerWiggle();
          setIsLoading(false);
          return;
        }
        if (password !== confirmPassword) {
          setError('Şifreler birbiriyle eşleşmiyor!');
          triggerWiggle();
          setIsLoading(false);
          return;
        }

        const { salt, hash } = await setupPassword(password);
        localStorage.setItem('muhasebe_email', email.trim().toLowerCase());
        localStorage.setItem('muhasebe_salt', salt);
        localStorage.setItem('muhasebe_hash', hash);
        
        // Derive key and unlock
        const derivedKey = await verifyPassword(password, salt, hash);
        if (derivedKey) {
          onUnlock(derivedKey);
        } else {
          setError('Şifre doğrulanırken hata oluştu!');
          triggerWiggle();
        }
      } else {
        // Unlock
        const savedEmail = localStorage.getItem('muhasebe_email') || '';
        if (email.trim().toLowerCase() !== savedEmail.trim().toLowerCase()) {
          setError('Girdiğiniz E-posta veya Şifre hatalı!');
          triggerWiggle();
          setIsLoading(false);
          return;
        }

        const salt = localStorage.getItem('muhasebe_salt') || '';
        const hash = localStorage.getItem('muhasebe_hash') || '';
        
        const derivedKey = await verifyPassword(password, salt, hash);
        if (derivedKey) {
          onUnlock(derivedKey);
        } else {
          setError('Girdiğiniz E-posta veya Şifre hatalı!');
          triggerWiggle();
        }
      }
    } catch (err: any) {
      console.error(err);
      setError('İşlem sırasında beklenmedik bir hata oluştu.');
      triggerWiggle();
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div style={{
      position: 'fixed',
      inset: 0,
      zIndex: 9999,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      background: 'radial-gradient(circle at center, #0f172a 0%, #020617 100%)',
      padding: '20px',
      color: '#f8fafc'
    }}>
      <div 
        className={`glass-card ${isWiggling ? 'animate-wiggle' : 'animate-fade-in'}`} 
        style={{
          width: '100%',
          maxWidth: '440px',
          padding: '40px 32px',
          backgroundColor: 'rgba(15, 23, 42, 0.45)',
          border: '1px solid rgba(255, 255, 255, 0.08)',
          boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.5)',
          borderRadius: '24px',
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          gap: '24px',
          textAlign: 'center',
          transition: 'transform 0.5s ease-in-out'
        }}
      >
        {/* Shield Icon Container */}
        <div style={{
          width: '72px',
          height: '72px',
          borderRadius: '50%',
          backgroundColor: isFirstTime ? 'rgba(16, 185, 129, 0.15)' : 'rgba(99, 102, 241, 0.15)',
          color: isFirstTime ? '#10b981' : '#6366f1',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          border: `1px solid ${isFirstTime ? 'rgba(16, 185, 129, 0.25)' : 'rgba(99, 102, 241, 0.25)'}`
        }}>
          {isFirstTime ? <ShieldCheck size={36} /> : <Lock size={32} />}
        </div>

        <div>
          <h2 className="text-h2" style={{ color: '#ffffff', fontSize: '1.4rem', fontWeight: 800, marginBottom: '8px' }}>
            {isFirstTime ? 'Yönetici Hesabı Oluşturun' : 'Muhasebe Giriş Portalı'}
          </h2>
          <p style={{ color: '#94a3b8', fontSize: '0.85rem', lineHeight: '145%', maxWidth: '340px', margin: '0 auto' }}>
            {isFirstTime 
              ? 'Verilerinizi yerel bilgisayarınızda AES-256 ile şifrelemek için e-posta ve şifre belirleyin.'
              : 'Verilerin şifresini çözmek ve sisteme erişmek için giriş yapın.'
            }
          </p>
        </div>

        <form onSubmit={handleAction} style={{ width: '100%', display: 'flex', flexDirection: 'column', gap: '16px' }}>
          {/* Email Input */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', textAlign: 'left' }}>
            <label style={{ fontSize: '0.75rem', fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
              E-posta Adresi
            </label>
            <div style={{ position: 'relative' }}>
              <Mail size={16} style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)', color: '#64748b' }} />
              <input
                type="email"
                placeholder="muhasebe@vantso.org.tr"
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="form-input"
                style={{
                  width: '100%',
                  paddingLeft: '38px',
                  paddingRight: '12px',
                  backgroundColor: 'rgba(15, 23, 42, 0.6)',
                  border: '1px solid rgba(255, 255, 255, 0.1)',
                  color: '#ffffff',
                  borderRadius: '10px',
                  height: '42px',
                  fontSize: '0.9rem'
                }}
              />
            </div>
          </div>

          {/* Password Input */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', textAlign: 'left' }}>
            <label style={{ fontSize: '0.75rem', fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
              {isFirstTime ? 'Şifre Belirleyin' : 'Şifre'}
            </label>
            <div style={{ position: 'relative' }}>
              <KeyRound size={16} style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)', color: '#64748b' }} />
              <input
                type={showPassword ? 'text' : 'password'}
                placeholder="••••••••"
                required
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                className="form-input"
                style={{
                  width: '100%',
                  paddingLeft: '38px',
                  paddingRight: '38px',
                  backgroundColor: 'rgba(15, 23, 42, 0.6)',
                  border: '1px solid rgba(255, 255, 255, 0.1)',
                  color: '#ffffff',
                  borderRadius: '10px',
                  height: '42px',
                  fontSize: '0.9rem'
                }}
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                style={{
                  position: 'absolute',
                  right: '12px',
                  top: '50%',
                  transform: 'translateY(-50%)',
                  background: 'none',
                  border: 'none',
                  color: '#64748b',
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  padding: 0
                }}
              >
                {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
              </button>
            </div>
          </div>

          {/* Confirm Password (only first time) */}
          {isFirstTime && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', textAlign: 'left' }}>
              <label style={{ fontSize: '0.75rem', fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                Şifreyi Onaylayın
              </label>
              <div style={{ position: 'relative' }}>
                <KeyRound size={16} style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)', color: '#64748b' }} />
                <input
                  type={showPassword ? 'text' : 'password'}
                  placeholder="••••••••"
                  required
                  value={confirmPassword}
                  onChange={(e) => setConfirmPassword(e.target.value)}
                  className="form-input"
                  style={{
                    width: '100%',
                    paddingLeft: '38px',
                    paddingRight: '38px',
                    backgroundColor: 'rgba(15, 23, 42, 0.6)',
                    border: '1px solid rgba(255, 255, 255, 0.1)',
                    color: '#ffffff',
                    borderRadius: '10px',
                    height: '42px',
                    fontSize: '0.9rem'
                  }}
                />
              </div>
            </div>
          )}

          {error && (
            <div style={{
              fontSize: '0.8rem',
              color: '#f87171',
              backgroundColor: 'rgba(239, 68, 68, 0.1)',
              border: '1px solid rgba(239, 68, 68, 0.2)',
              padding: '8px 12px',
              borderRadius: '8px',
              textAlign: 'left'
            }}>
              ⚠️ {error}
            </div>
          )}

          <button
            type="submit"
            disabled={isLoading}
            className="btn btn-primary"
            style={{
              width: '100%',
              height: '42px',
              justifyContent: 'center',
              backgroundColor: isFirstTime ? '#10b981' : '#6366f1',
              borderRadius: '10px',
              fontSize: '0.9rem',
              fontWeight: 700,
              color: '#ffffff',
              boxShadow: 'none',
              marginTop: '8px'
            }}
          >
            {isLoading ? 'İşleniyor...' : isFirstTime ? 'Hesap Oluştur ve Kilitle' : 'Giriş Yap'}
          </button>
        </form>

        <div style={{ borderTop: '1px solid rgba(255, 255, 255, 0.08)', width: '100%', paddingTop: '16px' }}>
          <p style={{ fontSize: '0.7rem', color: '#64748b', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
            🔒 AES-256 End-to-End Local Encryption Active
          </p>
        </div>
      </div>

      <style>{`
        .animate-wiggle {
          animation: wiggle 0.5s ease-in-out;
        }
        @keyframes wiggle {
          0%, 100% { transform: translateX(0); }
          15%, 45%, 75% { transform: translateX(-6px); }
          30%, 60%, 90% { transform: translateX(6px); }
        }
      `}</style>
    </div>
  );
};
