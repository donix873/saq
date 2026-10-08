import {
  ArrowDown,
  BrainCircuit,
  CheckCircle2,
  Languages,
  LockKeyhole,
  MousePointerClick,
  Route,
  SearchCheck,
  ShieldCheck,
  Sparkles,
  X,
} from 'lucide-react'
import { useEffect, useRef, useState } from 'react'
import { AnalysisResultCard } from './components/AnalysisResultCard'
import { Header } from './components/Header'
import { History } from './components/History'
import { Scanner } from './components/Scanner'
import { translate } from './i18n'
import { analyzeInput } from './lib/analyzer'
import { clearHistory, loadHistory, saveHistory } from './lib/storage'
import type { AnalysisResult, Locale, ScanMode } from './types'

const LOCALE_KEY = 'saq.locale'
const SAVE_HISTORY_KEY = 'saq.saveHistory'

function initialLocale(): Locale {
  try {
    const stored = window.localStorage.getItem(LOCALE_KEY)
    return stored === 'kk' ? 'kk' : 'ru'
  } catch {
    return 'ru'
  }
}

function initialSavePreference(): boolean {
  try {
    return window.localStorage.getItem(SAVE_HISTORY_KEY) !== 'false'
  } catch {
    return true
  }
}

export default function App() {
  const [locale, setLocale] = useState<Locale>(initialLocale)
  const [mode, setMode] = useState<ScanMode>('message')
  const [input, setInput] = useState('')
  const [result, setResult] = useState<AnalysisResult | null>(null)
  const [history, setHistory] = useState<AnalysisResult[]>(loadHistory)
  const [saveToHistory, setSaveToHistory] = useState(initialSavePreference)
  const [isAnalyzing, setIsAnalyzing] = useState(false)
  const [error, setError] = useState('')
  const [toast, setToast] = useState('')
  const toastTimer = useRef<number | null>(null)

  const t = (key: string) => translate(locale, key)

  useEffect(() => {
    document.documentElement.lang = locale === 'kk' ? 'kk' : 'ru'
    try {
      window.localStorage.setItem(LOCALE_KEY, locale)
    } catch {
      // The interface still works when storage is unavailable.
    }
  }, [locale])

  useEffect(() => {
    try {
      window.localStorage.setItem(SAVE_HISTORY_KEY, String(saveToHistory))
    } catch {
      // Saving this preference is optional.
    }
  }, [saveToHistory])

  useEffect(
    () => () => {
      if (toastTimer.current) window.clearTimeout(toastTimer.current)
    },
    [],
  )

  const notify = (message: string) => {
    setToast(message)
    if (toastTimer.current) window.clearTimeout(toastTimer.current)
    toastTimer.current = window.setTimeout(() => setToast(''), 2800)
  }

  const runAnalysis = async () => {
    setError('')
    if (input.trim().length < 4) {
      setError(t('scanner.tooShort'))
      return
    }

    setIsAnalyzing(true)
    await new Promise((resolve) => window.setTimeout(resolve, 420))

    try {
      const nextResult = analyzeInput(input, mode)
      setResult(nextResult)

      if (saveToHistory) {
        const nextHistory = [nextResult, ...history.filter((item) => item.input !== nextResult.input)].slice(0, 20)
        setHistory(nextHistory)
        if (!saveHistory(nextHistory)) notify(t('history.storageError'))
      }

      window.setTimeout(() => {
        document.querySelector('#scan-result')?.scrollIntoView({ behavior: 'smooth', block: 'nearest' })
      }, 50)
    } catch (analysisError) {
      setError(
        analysisError instanceof Error && analysisError.message === 'INPUT_TOO_SHORT'
          ? t('scanner.tooShort')
          : t('scanner.genericError'),
      )
    } finally {
      setIsAnalyzing(false)
    }
  }

  const openHistoryItem = (item: AnalysisResult) => {
    setResult(item)
    setInput(item.input)
    setMode(item.mode)
    document.querySelector('#scanner')?.scrollIntoView({ behavior: 'smooth' })
  }

  const deleteHistoryItem = (id: string) => {
    const nextHistory = history.filter((item) => item.id !== id)
    setHistory(nextHistory)
    saveHistory(nextHistory)
  }

  const deleteAllHistory = () => {
    if (!window.confirm(t('history.confirmClear'))) return
    setHistory([])
    clearHistory()
  }

  return (
    <div id="top" className="min-h-screen overflow-hidden bg-[#071017] text-slate-100">
      <div className="app-background" aria-hidden="true" />
      <Header locale={locale} onLocaleChange={setLocale} t={t} />

      <main className="relative z-10">
        <section className="mx-auto grid max-w-7xl gap-12 px-4 pt-16 pb-12 sm:px-6 sm:pt-24 lg:grid-cols-[1.02fr_.98fr] lg:items-center lg:px-8 lg:pt-28 lg:pb-20">
          <div>
            <div className="mb-6 inline-flex items-center gap-2 rounded-full border border-cyan-300/20 bg-cyan-300/7 px-3 py-2 text-[10px] font-bold tracking-[0.18em] text-cyan-200">
              <Sparkles size={14} />
              {t('hero.badge')}
            </div>
            <h1 className="max-w-4xl text-[clamp(2.8rem,7vw,5.9rem)] leading-[0.94] font-black tracking-[-0.055em] text-white">
              {t('hero.title.before')}{' '}
              <span className="text-gradient">{t('hero.title.accent')}</span>{' '}
              {t('hero.title.after')}
            </h1>
            <p className="mt-7 max-w-2xl text-base leading-7 text-slate-400 sm:text-lg sm:leading-8">
              {t('hero.subtitle')}
            </p>

            <div className="mt-8 flex flex-wrap gap-3">
              <a href="#scanner" className="primary-button px-5!">
                <SearchCheck size={18} />
                {t('nav.scanner')}
              </a>
              <a href="#how" className="secondary-button px-5!">
                {t('nav.how')}
                <ArrowDown size={16} />
              </a>
            </div>

            <div className="mt-10 grid max-w-2xl grid-cols-1 gap-3 sm:grid-cols-3">
              {[
                [LockKeyhole, 'hero.local'],
                [BrainCircuit, 'hero.explainable'],
                [Languages, 'hero.bilingual'],
              ].map(([Icon, key]) => {
                const FeatureIcon = Icon as typeof LockKeyhole
                return (
                  <div key={key as string} className="flex items-center gap-3 rounded-2xl border border-white/7 bg-white/[0.025] p-3.5">
                    <FeatureIcon size={18} className="shrink-0 text-cyan-300" />
                    <span className="text-xs font-semibold text-slate-300">{t(key as string)}</span>
                  </div>
                )
              })}
            </div>
          </div>

          <div className="hero-visual relative mx-auto hidden h-[31rem] w-full max-w-[34rem] lg:block">
            <div className="absolute inset-8 rounded-full border border-cyan-300/10" />
            <div className="absolute inset-20 rounded-full border border-cyan-300/15" />
            <div className="absolute inset-32 rounded-full border border-cyan-300/20" />
            <div className="absolute inset-[7rem] grid place-items-center">
              <div className="absolute size-64 rounded-full bg-cyan-300/10 blur-3xl" />
              <img src="/saq-logo.png" alt="SAQ" className="relative z-10 w-64 drop-shadow-[0_0_45px_rgba(73,239,234,.28)]" />
            </div>
            <div className="absolute top-[12%] right-[8%] flex items-center gap-2 rounded-2xl border border-red-300/15 bg-red-300/7 px-3 py-2 text-xs text-red-200 backdrop-blur">
              <span className="size-2 animate-pulse rounded-full bg-red-400" />
              Social engineering
            </div>
            <div className="absolute bottom-[16%] left-[4%] flex items-center gap-2 rounded-2xl border border-emerald-300/15 bg-emerald-300/7 px-3 py-2 text-xs text-emerald-200 backdrop-blur">
              <ShieldCheck size={15} />
              Explainable protection
            </div>
          </div>
        </section>

        <section id="scanner" className="scroll-mt-24 border-y border-white/6 bg-white/[0.015]">
          <div className="mx-auto grid max-w-7xl gap-5 px-4 py-12 sm:px-6 sm:py-16 lg:grid-cols-2 lg:items-start lg:px-8">
            <Scanner
              locale={locale}
              mode={mode}
              input={input}
              isAnalyzing={isAnalyzing}
              saveToHistory={saveToHistory}
              error={error}
              onModeChange={(nextMode) => {
                setMode(nextMode)
                setError('')
              }}
              onInputChange={(value) => {
                setInput(value)
                setError('')
              }}
              onAnalyze={runAnalysis}
              onSaveToHistoryChange={setSaveToHistory}
              onNotify={notify}
              t={t}
            />
            <div id="scan-result" className="scroll-mt-24">
              <AnalysisResultCard result={result} onNotify={notify} t={t} />
            </div>
          </div>
        </section>

        <section id="how" className="section-shell scroll-mt-24">
          <p className="section-eyebrow">
            <Route size={15} />
            {t('how.eyebrow')}
          </p>
          <h2 className="section-title max-w-3xl">{t('how.title')}</h2>

          <div className="mt-10 grid gap-4 md:grid-cols-3">
            {[
              [MousePointerClick, '01', 'how.step1.title', 'how.step1.text', 'text-cyan-300'],
              [BrainCircuit, '02', 'how.step2.title', 'how.step2.text', 'text-orange-300'],
              [CheckCircle2, '03', 'how.step3.title', 'how.step3.text', 'text-emerald-300'],
            ].map(([Icon, number, title, copy, color]) => {
              const StepIcon = Icon as typeof MousePointerClick
              return (
                <article key={number as string} className="glass-panel rounded-3xl border border-white/8 p-6 sm:p-7">
                  <div className="flex items-center justify-between">
                    <StepIcon size={24} className={color as string} />
                    <span className="text-xs font-black tracking-[0.18em] text-slate-600">{number as string}</span>
                  </div>
                  <h3 className="mt-8 text-xl font-bold text-white">{t(title as string)}</h3>
                  <p className="mt-3 leading-7 text-slate-400">{t(copy as string)}</p>
                </article>
              )
            })}
          </div>
        </section>

        <section id="privacy" className="section-shell scroll-mt-24 pt-0!">
          <div className="relative overflow-hidden rounded-[2rem] border border-emerald-300/15 bg-emerald-300/[0.035] p-6 sm:p-10 lg:p-12">
            <div className="absolute -top-32 -right-32 size-80 rounded-full bg-emerald-300/10 blur-3xl" aria-hidden="true" />
            <div className="relative grid gap-10 lg:grid-cols-[.85fr_1.15fr] lg:items-center">
              <div>
                <p className="section-eyebrow text-emerald-300!">
                  <LockKeyhole size={15} />
                  {t('privacy.eyebrow')}
                </p>
                <h2 className="section-title">{t('privacy.title')}</h2>
                <p className="mt-5 max-w-xl leading-7 text-slate-400">{t('privacy.text')}</p>
              </div>
              <div className="grid gap-3">
                {['privacy.point1', 'privacy.point2', 'privacy.point3'].map((key) => (
                  <div key={key} className="flex items-center gap-3 rounded-2xl border border-emerald-300/10 bg-[#071017]/45 p-4">
                    <span className="grid size-8 shrink-0 place-items-center rounded-full bg-emerald-300 text-slate-950">
                      <CheckCircle2 size={17} />
                    </span>
                    <span className="font-semibold text-slate-200">{t(key)}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </section>

        <History
          history={history}
          onOpen={openHistoryItem}
          onDelete={deleteHistoryItem}
          onClear={deleteAllHistory}
          t={t}
        />
      </main>

      <footer className="relative z-10 border-t border-white/7 px-4 py-8 sm:px-6 lg:px-8">
        <div className="mx-auto flex max-w-7xl flex-col justify-between gap-4 sm:flex-row sm:items-center">
          <div className="flex items-center gap-3">
            <img src="/saq-logo.png" alt="" className="size-9 object-contain" />
            <div>
              <p className="font-black tracking-[0.16em] text-white">SAQ</p>
              <p className="text-xs text-cyan-300">{t('footer.tagline')}</p>
            </div>
          </div>
          <p className="text-xs text-slate-600">© {new Date().getFullYear()} SAQ · {t('footer.made')}</p>
        </div>
      </footer>

      {toast && (
        <div className="fixed right-4 bottom-4 z-[100] flex max-w-sm items-center gap-3 rounded-2xl border border-cyan-300/20 bg-[#10232c]/95 px-4 py-3 text-sm font-semibold text-slate-100 shadow-2xl backdrop-blur-xl">
          <CheckCircle2 size={17} className="shrink-0 text-cyan-300" />
          <span>{toast}</span>
          <button type="button" onClick={() => setToast('')} className="ml-2 text-slate-500 hover:text-white" aria-label="Close">
            <X size={15} />
          </button>
        </div>
      )}
    </div>
  )
}
