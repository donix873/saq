import {
  Clipboard,
  Eraser,
  Link2,
  LoaderCircle,
  LockKeyhole,
  MessageSquareText,
  ScanSearch,
  Sparkles,
} from 'lucide-react'
import { useId } from 'react'
import type { Locale, ScanMode } from '../types'

const examples: Record<Locale, Array<{ key: string; mode: ScanMode; value: string }>> = {
  ru: [
    {
      key: 'example.bank',
      mode: 'message',
      value:
        'Здравствуйте. Служба безопасности Kaspi обнаружила подозрительную операцию. Срочно подтвердите данные по ссылке http://kaspi-secure-pay.cc/auth, иначе счёт будет заблокирован. Никому не сообщайте об этом.',
    },
    {
      key: 'example.job',
      mode: 'message',
      value:
        'Работа без опыта: ставьте лайки и получайте 50 000 тенге в день. Для активации переведите небольшой депозит прямо сейчас.',
    },
    {
      key: 'example.invest',
      mode: 'message',
      value:
        'Гарантированный доход без риска. Мы удвоим вложение за неделю. Последний шанс — внесите депозит сегодня.',
    },
    {
      key: 'example.safe',
      mode: 'message',
      value: 'Привет! Встреча завтра в 15:00 в школьной библиотеке. Если планы изменятся, напиши мне.',
    },
  ],
  kk: [
    {
      key: 'example.bank',
      mode: 'message',
      value:
        'Сәлеметсіз бе. Kaspi қорғау қызметі күмәнді операцияны анықтады. Шотыңыз бұғатталмауы үшін http://kaspi-secure-pay.cc/auth сілтемесі арқылы деректерді шұғыл растаңыз. Ешкімге айтпаңыз.',
    },
    {
      key: 'example.job',
      mode: 'message',
      value:
        'Тәжірибесіз жұмыс: лайк басып күніне 50 000 теңге табыңыз. Белсендіру үшін қазір шағын депозит салу керек.',
    },
    {
      key: 'example.invest',
      mode: 'message',
      value:
        'Тәуекелсіз кепілдендірілген табыс. Ақшаңызды бір аптада екі есе көбейтеміз. Соңғы мүмкіндік — бүгін депозит салыңыз.',
    },
    {
      key: 'example.safe',
      mode: 'message',
      value: 'Сәлем! Ертең сағат 15:00-де мектеп кітапханасында кездесеміз. Жоспар өзгерсе, маған жаз.',
    },
  ],
}

interface ScannerProps {
  locale: Locale
  mode: ScanMode
  input: string
  isAnalyzing: boolean
  saveToHistory: boolean
  error: string
  onModeChange: (mode: ScanMode) => void
  onInputChange: (value: string) => void
  onAnalyze: () => void
  onSaveToHistoryChange: (value: boolean) => void
  onNotify: (message: string) => void
  t: (key: string) => string
}

export function Scanner({
  locale,
  mode,
  input,
  isAnalyzing,
  saveToHistory,
  error,
  onModeChange,
  onInputChange,
  onAnalyze,
  onSaveToHistoryChange,
  onNotify,
  t,
}: ScannerProps) {
  const inputId = useId()

  const paste = async () => {
    try {
      const value = await navigator.clipboard.readText()
      if (!value.trim()) {
        onNotify(t('scanner.clipboardEmpty'))
        return
      }
      onInputChange(value)
    } catch {
      onNotify(t('scanner.clipboardError'))
    }
  }

  return (
    <div className="glass-panel overflow-hidden rounded-[1.75rem] border border-white/10">
      <div className="border-b border-white/8 p-5 sm:p-7">
        <div className="flex flex-wrap items-start justify-between gap-4">
          <div>
            <div className="mb-2 flex items-center gap-2 text-[11px] font-bold tracking-[0.18em] text-cyan-300">
              <Sparkles size={14} />
              {t('scanner.eyebrow')}
            </div>
            <h2 className="text-xl font-bold text-white sm:text-2xl">{t('scanner.title')}</h2>
          </div>
          <div className="flex rounded-xl border border-white/10 bg-slate-950/45 p-1">
            {(['message', 'link'] as const).map((item) => {
              const Icon = item === 'message' ? MessageSquareText : Link2
              return (
                <button
                  key={item}
                  type="button"
                  onClick={() => onModeChange(item)}
                  className={`flex items-center gap-2 rounded-lg px-3 py-2 text-xs font-bold transition sm:text-sm ${
                    mode === item ? 'bg-white/10 text-white' : 'text-slate-400 hover:text-white'
                  }`}
                  aria-pressed={mode === item}
                >
                  <Icon size={15} />
                  {t(`scanner.${item}`)}
                </button>
              )
            })}
          </div>
        </div>
      </div>

      <div className="p-5 sm:p-7">
        <div className="relative">
          <label htmlFor={inputId} className="sr-only">
            {t(`scanner.${mode}`)}
          </label>
          <textarea
            id={inputId}
            value={input}
            onChange={(event) => onInputChange(event.target.value.slice(0, 3000))}
            placeholder={t(`scanner.${mode}Placeholder`)}
            rows={8}
            spellCheck
            className={`min-h-48 w-full resize-y rounded-2xl border bg-[#08141c]/80 px-4 py-4 pr-4 text-[15px] leading-7 text-slate-100 outline-none transition placeholder:text-slate-600 focus:ring-3 sm:px-5 ${
              error
                ? 'border-red-400/60 focus:border-red-400 focus:ring-red-400/10'
                : 'border-white/10 focus:border-cyan-300/55 focus:ring-cyan-300/10'
            }`}
          />
          <div className="pointer-events-none absolute right-3 bottom-3 rounded-lg bg-[#071017]/85 px-2 py-1 text-[10px] font-medium text-slate-500">
            {input.length}/3000
          </div>
        </div>

        {error && <p className="mt-2 text-sm text-red-300">{error}</p>}

        <div className="mt-3 flex flex-wrap items-center justify-between gap-3">
          <div className="flex gap-2">
            <button
              type="button"
              onClick={paste}
              className="secondary-button"
            >
              <Clipboard size={15} />
              {t('scanner.paste')}
            </button>
            {input && (
              <button type="button" onClick={() => onInputChange('')} className="secondary-button">
                <Eraser size={15} />
                {t('scanner.clear')}
              </button>
            )}
          </div>
          <div className="flex items-center gap-2 text-xs text-emerald-300/85">
            <LockKeyhole size={14} />
            {t('scanner.localNote')}
          </div>
        </div>

        <div className="mt-6">
          <p className="mb-3 text-[11px] font-bold tracking-[0.14em] text-slate-500 uppercase">
            {t('scanner.examples')}
          </p>
          <div className="flex flex-wrap gap-2">
            {examples[locale].map((example) => (
              <button
                key={example.key}
                type="button"
                onClick={() => {
                  onModeChange(example.mode)
                  onInputChange(example.value)
                }}
                className="rounded-full border border-white/9 bg-white/4 px-3 py-2 text-xs font-medium text-slate-400 transition hover:border-cyan-300/30 hover:bg-cyan-300/7 hover:text-cyan-100"
              >
                {t(example.key)}
              </button>
            ))}
          </div>
        </div>

        <label className="mt-6 flex cursor-pointer items-center gap-3 rounded-xl border border-white/7 bg-white/[0.025] p-3.5 text-sm text-slate-400">
          <input
            type="checkbox"
            checked={saveToHistory}
            onChange={(event) => onSaveToHistoryChange(event.target.checked)}
            className="size-4 accent-cyan-300"
          />
          {t('scanner.saveHistory')}
        </label>

        <button
          type="button"
          onClick={onAnalyze}
          disabled={isAnalyzing || input.trim().length < 4}
          className="primary-button mt-4 w-full"
        >
          {isAnalyzing ? (
            <LoaderCircle className="animate-spin" size={19} />
          ) : (
            <ScanSearch size={19} />
          )}
          {isAnalyzing ? t('scanner.analyzing') : t('scanner.analyze')}
        </button>
      </div>
    </div>
  )
}
