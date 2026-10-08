import { Menu, ShieldCheck, X } from 'lucide-react'
import { useState } from 'react'
import type { Locale } from '../types'

interface HeaderProps {
  locale: Locale
  onLocaleChange: (locale: Locale) => void
  t: (key: string) => string
}

export function Header({ locale, onLocaleChange, t }: HeaderProps) {
  const [open, setOpen] = useState(false)

  const nav = [
    { href: '#scanner', label: t('nav.scanner') },
    { href: '#how', label: t('nav.how') },
    { href: '#privacy', label: t('nav.privacy') },
    { href: '#history', label: t('nav.history') },
  ]

  return (
    <header className="sticky top-0 z-50 border-b border-white/7 bg-[#071017]/80 backdrop-blur-xl">
      <div className="mx-auto flex h-18 max-w-7xl items-center justify-between px-4 sm:px-6 lg:px-8">
        <a href="#top" className="group flex items-center gap-3" aria-label="SAQ">
          <span className="relative grid size-10 place-items-center overflow-hidden rounded-xl border border-cyan-300/25 bg-cyan-300/8">
            <img
              src="/saq-logo.png"
              alt=""
              className="size-9 object-contain transition-transform duration-300 group-hover:scale-105"
            />
          </span>
          <span>
            <span className="block text-lg font-black tracking-[0.18em] text-white">SAQ</span>
            <span className="hidden text-[9px] font-semibold tracking-[0.16em] text-cyan-200/65 sm:block">
              THINK BEFORE YOU TRUST
            </span>
          </span>
        </a>

        <nav className="hidden items-center gap-7 lg:flex" aria-label="Main navigation">
          {nav.map((item) => (
            <a
              key={item.href}
              href={item.href}
              className="text-sm font-medium text-slate-300 transition-colors hover:text-cyan-300"
            >
              {item.label}
            </a>
          ))}
        </nav>

        <div className="flex items-center gap-2">
          <div className="flex rounded-xl border border-white/10 bg-white/4 p-1" aria-label="Language">
            {(['ru', 'kk'] as const).map((item) => (
              <button
                key={item}
                type="button"
                onClick={() => onLocaleChange(item)}
                className={`rounded-lg px-2.5 py-1.5 text-xs font-bold transition ${
                  locale === item
                    ? 'bg-cyan-300 text-slate-950 shadow-[0_0_24px_rgba(73,239,234,.18)]'
                    : 'text-slate-400 hover:text-white'
                }`}
                aria-pressed={locale === item}
              >
                {item === 'ru' ? 'RU' : 'ҚАЗ'}
              </button>
            ))}
          </div>
          <button
            type="button"
            className="grid size-10 place-items-center rounded-xl border border-white/10 text-slate-300 hover:bg-white/5 lg:hidden"
            onClick={() => setOpen((value) => !value)}
            aria-label="Toggle menu"
            aria-expanded={open}
          >
            {open ? <X size={19} /> : <Menu size={19} />}
          </button>
        </div>
      </div>

      {open && (
        <nav className="border-t border-white/7 bg-[#071017] px-4 py-4 lg:hidden" aria-label="Mobile navigation">
          <div className="mx-auto grid max-w-7xl gap-2">
            {nav.map((item) => (
              <a
                key={item.href}
                href={item.href}
                onClick={() => setOpen(false)}
                className="flex items-center gap-3 rounded-xl px-3 py-3 text-sm font-medium text-slate-200 hover:bg-white/5"
              >
                <ShieldCheck size={16} className="text-cyan-300" />
                {item.label}
              </a>
            ))}
          </div>
        </nav>
      )}
    </header>
  )
}
