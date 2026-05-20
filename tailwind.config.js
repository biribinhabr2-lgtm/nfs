/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,jsx}'],
  safelist: [
    'bg-emerald-500/5','bg-emerald-500/10','border-emerald-500/20','text-emerald-400','text-emerald-300',
    'bg-rose-500/5','bg-rose-500/10','border-rose-500/20','text-rose-400','text-rose-300',
    'bg-amber-500/5','bg-amber-500/10','border-amber-500/20','text-amber-400','text-amber-300',
    'bg-indigo-500/5','bg-indigo-500/10','border-indigo-500/20','text-indigo-400','text-indigo-300',
    'bg-violet-500/5','bg-violet-500/10','bg-violet-500/20','border-violet-500/20','text-violet-400','text-violet-300',
    'bg-sky-500/5','bg-sky-500/10','border-sky-500/20','text-sky-400','text-sky-300',
  ],
  theme: {
    extend: {
      fontFamily: {
        sans: ['"DM Sans"', 'system-ui', 'sans-serif'],
        mono: ['"DM Mono"', 'monospace'],
      },
    },
  },
  plugins: [],
}
