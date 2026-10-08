import {
  AlertTriangle,
  Check,
  CheckCircle2,
  ClipboardCopy,
  Info,
  Radar,
  ShieldAlert,
  Sparkles,
} from 'lucide-react'
import type { AnalysisResult, RiskLevel } from '../types'

const riskStyles: Record<
  RiskLevel,
  { accent: string; badge: string; ring: string; icon: typeof Info }
> = {
  low: {
    accent: 'text-emerald-300',
    badge: 'border-emerald-300/30 bg-emerald-300/10 text-emerald-200',
    ring: '#5ce3a8',
    icon: CheckCircle2,
  },
  medium: {
    accent: 'text-amber-300',
    badge: 'border-amber-300/30 bg-amber-300/10 text-amber-200',
    ring: '#ffd166',
    icon: AlertTriangle,
  },
  high: {
    accent: 'text-orange-300',
    badge: 'border-orange-300/30 bg-orange-300/10 text-orange-200',
    ring: '#ff9b63',
    icon: ShieldAlert,
  },
  critical: {
    accent: 'text-red-300',
    badge: 'border-red-300/30 bg-red-300/10 text-red-200',
    ring: '#ff6b6b',
    icon: ShieldAlert,
  },
}

interface AnalysisResultCardProps {
  result: AnalysisResult | null
  onNotify: (message: string) => void
  t: (key: string) => string
}

export function AnalysisResultCard({ result, onNotify, t }: AnalysisResultCardProps) {
  if (!result) {
    return (
      <div className="glass-panel grid min-h-[36rem] place-items-center rounded-[1.75rem] border border-white/10 p-8 text-center">
        <div className="max-w-sm">
          <div className="radar-pulse mx-auto mb-7 grid size-24 place-items-center rounded-full border border-cyan-300/20 bg-cyan-300/6">
            <Radar size={38} className="text-cyan-300" />
          </div>
          <h2 className="text-2xl font-bold text-white">{t('result.waitingTitle')}</h2>
          <p className="mt-3 leading-7 text-slate-400">{t('result.waitingText')}</p>
        </div>
      </div>
    )
  }

  const style = riskStyles[result.level]
  const Icon = style.icon

  const copyReport = async () => {
    const lines = [
      `SAQ — ${t('result.risk')}: ${t(`risk.${result.level}`)} (${result.score}/100)`,
      '',
      t('result.detected'),
      ...(result.signals.length
        ? result.signals.map((signal) => `• ${t(`signal.${signal.code}.title`)}`)
        : [`• ${t('result.noSignals')}`]),
      '',
      t('result.actions'),
      ...result.actions.map((action) => `• ${t(`action.${action}`)}`),
      '',
      t('result.disclaimer'),
    ]

    try {
      await navigator.clipboard.writeText(lines.join('\n'))
      onNotify(t('result.copied'))
    } catch {
      onNotify(t('scanner.clipboardError'))
    }
  }

  return (
    <div className="result-enter glass-panel overflow-hidden rounded-[1.75rem] border border-white/10">
      <div className="border-b border-white/8 p-5 sm:p-7">
        <div className="flex flex-col gap-5 sm:flex-row sm:items-center">
          <div
            className="relative grid size-28 shrink-0 place-items-center rounded-full"
            style={{
              background: `conic-gradient(${style.ring} ${result.score * 3.6}deg, rgba(255,255,255,.07) 0deg)`,
            }}
          >
            <div className="grid size-[6.2rem] place-items-center rounded-full bg-[#0a161e]">
              <div className="text-center">
                <strong className={`block text-3xl font-black ${style.accent}`}>{result.score}</strong>
                <span className="text-[10px] font-bold tracking-wider text-slate-500">/ 100</span>
              </div>
            </div>
          </div>

          <div className="min-w-0 flex-1">
            <p className="text-[11px] font-bold tracking-[0.16em] text-slate-500 uppercase">
              {t('result.risk')}
            </p>
            <div className="mt-2 flex flex-wrap items-center gap-3">
              <span className={`inline-flex items-center gap-2 rounded-full border px-3 py-1.5 text-sm font-bold ${style.badge}`}>
                <Icon size={15} />
                {t(`risk.${result.level}`)}
              </span>
              <span className="rounded-full border border-white/8 bg-white/4 px-3 py-1.5 text-xs text-slate-400">
                {t(`mode.${result.mode}`)}
              </span>
            </div>
            <p className="mt-3 leading-6 text-slate-300">{t(`risk.${result.level}Text`)}</p>
          </div>
        </div>
      </div>

      <div className="space-y-7 p-5 sm:p-7">
        <section>
          <h3 className="mb-3 flex items-center gap-2 text-sm font-bold tracking-wide text-white">
            <Sparkles size={16} className="text-cyan-300" />
            {t('result.detected')}
          </h3>
          {result.signals.length ? (
            <div className="space-y-2.5">
              {result.signals.map((signal) => (
                <div key={signal.code} className="rounded-xl border border-white/7 bg-white/[0.025] p-3.5">
                  <div className="flex items-start gap-3">
                    <span className="mt-1 size-2 shrink-0 rounded-full bg-orange-400 shadow-[0_0_14px_rgba(251,146,60,.55)]" />
                    <div>
                      <div className="flex flex-wrap items-center gap-2">
                        <p className="text-sm font-bold text-slate-100">{t(`signal.${signal.code}.title`)}</p>
                        <code className="rounded bg-white/5 px-1.5 py-0.5 text-[10px] text-slate-500">
                          {signal.evidence}
                        </code>
                      </div>
                      <p className="mt-1 text-xs leading-5 text-slate-400">
                        {t(`signal.${signal.code}.description`)}
                      </p>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <div className="flex items-center gap-3 rounded-xl border border-emerald-300/15 bg-emerald-300/5 p-4 text-sm text-emerald-100">
              <CheckCircle2 size={18} className="text-emerald-300" />
              {t('result.noSignals')}
            </div>
          )}
        </section>

        <section>
          <h3 className="mb-3 flex items-center gap-2 text-sm font-bold tracking-wide text-white">
            <Check size={16} className="text-emerald-300" />
            {t('result.actions')}
          </h3>
          <ol className="space-y-2.5">
            {result.actions.map((action, index) => (
              <li key={action} className="flex gap-3 rounded-xl border border-emerald-300/10 bg-emerald-300/[0.035] p-3.5">
                <span className="grid size-6 shrink-0 place-items-center rounded-full bg-emerald-300 text-[11px] font-black text-slate-950">
                  {index + 1}
                </span>
                <span className="text-sm leading-6 text-slate-200">{t(`action.${action}`)}</span>
              </li>
            ))}
          </ol>
        </section>

        <div className="flex flex-col gap-3 border-t border-white/8 pt-5 sm:flex-row sm:items-center sm:justify-between">
          <p className="max-w-sm text-xs leading-5 text-slate-500">{t('result.disclaimer')}</p>
          <button type="button" onClick={copyReport} className="secondary-button shrink-0">
            <ClipboardCopy size={15} />
            {t('result.copy')}
          </button>
        </div>
      </div>
    </div>
  )
}
