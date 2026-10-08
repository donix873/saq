import { Clock3, ExternalLink, HistoryIcon, Trash2 } from 'lucide-react'
import type { AnalysisResult, RiskLevel } from '../types'

const levelDot: Record<RiskLevel, string> = {
  low: 'bg-emerald-300',
  medium: 'bg-amber-300',
  high: 'bg-orange-300',
  critical: 'bg-red-300',
}

interface HistoryProps {
  history: AnalysisResult[]
  onOpen: (result: AnalysisResult) => void
  onDelete: (id: string) => void
  onClear: () => void
  t: (key: string) => string
}

export function History({ history, onOpen, onDelete, onClear, t }: HistoryProps) {
  return (
    <section id="history" className="section-shell scroll-mt-24">
      <div className="mb-8 flex flex-col justify-between gap-4 sm:flex-row sm:items-end">
        <div>
          <p className="section-eyebrow">
            <HistoryIcon size={15} />
            {t('history.eyebrow')}
          </p>
          <h2 className="section-title">{t('history.title')}</h2>
        </div>
        {history.length > 0 && (
          <button type="button" onClick={onClear} className="secondary-button self-start text-red-200 sm:self-auto">
            <Trash2 size={15} />
            {t('history.clear')}
          </button>
        )}
      </div>

      {history.length === 0 ? (
        <div className="glass-panel rounded-3xl border border-dashed border-white/10 p-10 text-center text-slate-500">
          <Clock3 className="mx-auto mb-4 text-slate-600" size={30} />
          {t('history.empty')}
        </div>
      ) : (
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          {history.map((item) => (
            <article key={item.id} className="glass-panel group rounded-2xl border border-white/8 p-4 transition hover:-translate-y-0.5 hover:border-white/15">
              <div className="flex items-center justify-between gap-3">
                <span className="flex items-center gap-2 text-xs font-bold text-slate-300">
                  <span className={`size-2 rounded-full ${levelDot[item.level]}`} />
                  {t(`risk.${item.level}`)} · {item.score}/100
                </span>
                <time className="text-[10px] text-slate-600">
                  {new Intl.DateTimeFormat(undefined, { dateStyle: 'short', timeStyle: 'short' }).format(
                    new Date(item.createdAt),
                  )}
                </time>
              </div>
              <p className="mt-3 line-clamp-3 min-h-15 text-sm leading-5 text-slate-400">{item.input}</p>
              <div className="mt-4 flex items-center justify-between border-t border-white/7 pt-3">
                <button
                  type="button"
                  onClick={() => onOpen(item)}
                  className="flex items-center gap-1.5 text-xs font-bold text-cyan-300 hover:text-cyan-200"
                >
                  <ExternalLink size={13} />
                  {t('history.open')}
                </button>
                <button
                  type="button"
                  onClick={() => onDelete(item.id)}
                  className="grid size-8 place-items-center rounded-lg text-slate-600 transition hover:bg-red-300/8 hover:text-red-300"
                  aria-label={t('history.delete')}
                >
                  <Trash2 size={14} />
                </button>
              </div>
            </article>
          ))}
        </div>
      )}
    </section>
  )
}
