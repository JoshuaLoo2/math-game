/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    './數學小遊戲.html',
    './js_*.html'
  ],
  safelist: [
    'bg-slate-100',
    'text-slate-300',
    'text-slate-500',
    'bg-emerald-100',
    'text-emerald-500',
    'text-emerald-700',
    'bg-amber-100',
    'text-amber-500',
    'text-amber-700',
    'bg-green-50/70',
    'border-green-200',
    'hover:border-green-400',
    'bg-gray-50/70',
    'border-gray-200',
    'opacity-60',
    'cursor-not-allowed',
    'bg-white/70',
    'border-emerald-200',
    'hover:border-emerald-400',
    'cursor-pointer',
    'text-blue-600',
    'text-emerald-600',
    'text-fuchsia-600',
    'text-amber-600',
    'drop-shadow-[0_8px_16px_rgba(148,163,184,0.24)]',
    'drop-shadow-[0_8px_18px_rgba(16,185,129,0.28)]',
    'drop-shadow-[0_8px_18px_rgba(245,158,11,0.26)]'
  ],
  theme: {
    extend: {}
  },
  plugins: []
};
