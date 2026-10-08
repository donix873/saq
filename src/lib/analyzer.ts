import type {
  ActionCode,
  AnalysisResult,
  RiskLevel,
  RiskSignal,
  ScanMode,
  SignalCode,
} from '../types'

interface PatternRule {
  code: SignalCode
  weight: number
  pattern: RegExp
}

const textRules: PatternRule[] = [
  {
    code: 'urgency',
    weight: 14,
    pattern:
      /(срочно|немедленно|прямо сейчас|в течение \d+|последний шанс|шұғыл|дереу|қазір|соңғы мүмкіндік|\d+ минут ішінде)/i,
  },
  {
    code: 'threat',
    weight: 17,
    pattern:
      /(заблокир(?:ован|уем|уют)|приостанов(?:лен|им)|штраф|уголовн|арест|потеряете доступ|бұғаттал(?:ады|ған)|айыппұл|қылмыстық|қолжетімділіктен айырыласыз)/i,
  },
  {
    code: 'impersonation',
    weight: 12,
    pattern:
      /(служба безопасности|сотрудник банка|полици(?:я|и)|налоговая|госуслуг|e-?gov|қорғау қызметі|банк қызметкері|полиция|мемлекеттік орган)/i,
  },
  {
    code: 'secretRequest',
    weight: 30,
    pattern:
      /(код из смс|код подтверждения|одноразов(?:ый|ого) код|otp|парол[ья]|cvv|cvc|пин[ -]?код|смс код|sms растау коды|растау коды|бір реттік код|құпия сөз|пин[ -]?код)/i,
  },
  {
    code: 'personalData',
    weight: 19,
    pattern:
      /(иин|жсн|номер удостоверения|данные карты|номер карты|паспортные данные|жеке куәлік|карта деректері|карта нөмірі)/i,
  },
  {
    code: 'paymentRequest',
    weight: 22,
    pattern:
      /(переведите|оплатите|предоплат|внесите депозит|безопасный сч[её]т|резервный сч[её]т|ақша аудар|төлем жаса|алдын ала төлем|депозит сал|қауіпсіз шот)/i,
  },
  {
    code: 'tooGoodToBeTrue',
    weight: 16,
    pattern:
      /(гарантированн(?:ый|ая) доход|без риска|вы выиграли|приз|удвоим|л[её]гкие деньги|кепілдендірілген табыс|тәуекелсіз|ұтып алдыңыз|сыйлық|ақшаны екі есе)/i,
  },
  {
    code: 'suspiciousJob',
    weight: 15,
    pattern:
      /(работа без опыта|заработок на лайках|оплата за лайки|вакансия.*предоплат|тәжірибесіз жұмыс|лайк үшін ақша|жұмыс.*алдын ала төлем)/i,
  },
  {
    code: 'isolation',
    weight: 18,
    pattern:
      /(никому не говорите|не кладите трубку|не звоните в банк|держите в секрете|ешкімге айтпаңыз|тұтқаны қоймаңыз|банкке қоңырау шалмаңыз|құпия сақтаңыз)/i,
  },
  {
    code: 'remoteAccess',
    weight: 27,
    pattern:
      /(anydesk|teamviewer|rustdesk|удал[её]нн(?:ый|ого) доступ|демонстраци(?:я|ю) экрана|қашықтан қол жеткізу|экранды көрсету)/i,
  },
]

const suspiciousTlds = new Set([
  'click',
  'top',
  'xyz',
  'site',
  'online',
  'work',
  'live',
  'cc',
  'info',
  'buzz',
])

const shorteners = new Set([
  'bit.ly',
  'tinyurl.com',
  't.co',
  'cutt.ly',
  'clck.ru',
  'rb.gy',
  'goo.gl',
])

const brands: Array<{ pattern: RegExp; domains: string[] }> = [
  { pattern: /kaspi/i, domains: ['kaspi.kz'] },
  { pattern: /halyk|халық банк|народный банк/i, domains: ['halykbank.kz'] },
  { pattern: /egov|е-?гов/i, domains: ['egov.kz'] },
  { pattern: /forte/i, domains: ['forte.kz'] },
  { pattern: /jusan/i, domains: ['jusan.kz'] },
]

function createId(): string {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) {
    return crypto.randomUUID()
  }

  return `${Date.now()}-${Math.random().toString(16).slice(2)}`
}

function extractUrls(input: string): string[] {
  const matches = input.match(
    /(?:https?:\/\/|www\.)[^\s<>"']+|\b[a-z\d][a-z\d-]*(?:\.[a-z\d-]+)+(?:\/[^\s<>"']*)?/gi,
  )

  if (!matches) return []

  return [...new Set(matches.map((url) => url.replace(/[),.;!?]+$/g, '')))]
}

function asUrl(value: string): URL | null {
  try {
    return new URL(/^https?:\/\//i.test(value) ? value : `https://${value}`)
  } catch {
    return null
  }
}

function levelFromScore(score: number): RiskLevel {
  if (score >= 75) return 'critical'
  if (score >= 50) return 'high'
  if (score >= 25) return 'medium'
  return 'low'
}

function actionList(signals: RiskSignal[]): ActionCode[] {
  const codes = new Set(signals.map((signal) => signal.code))
  const actions = new Set<ActionCode>()

  if (
    ['suspiciousTld', 'shortener', 'ipAddress', 'punycode', 'insecureProtocol', 'obfuscatedUrl', 'brandMismatch'].some(
      (code) => codes.has(code as SignalCode),
    )
  ) {
    actions.add('dontOpenLink')
  }
  if (codes.has('secretRequest')) actions.add('dontShareCodes')
  if (codes.has('personalData')) actions.add('dontShareData')
  if (codes.has('paymentRequest')) actions.add('stopPayment')
  if (codes.has('remoteAccess')) actions.add('dontInstallApps')
  if (codes.has('impersonation') || codes.has('brandMismatch')) actions.add('verifyOfficially')
  if (signals.length >= 3) actions.add('contactSupport')
  if (signals.length > 0) actions.add('preserveEvidence')
  if (actions.size === 0) actions.add('verifySender')

  return [...actions]
}

export function analyzeInput(rawInput: string, requestedMode?: ScanMode): AnalysisResult {
  const input = rawInput.trim()
  if (input.length < 4) {
    throw new Error('INPUT_TOO_SHORT')
  }

  const urls = extractUrls(input)
  const mode: ScanMode = requestedMode ?? (urls.length === 1 && input.length < 180 ? 'link' : 'message')
  const signals: RiskSignal[] = []
  const seen = new Set<SignalCode>()

  const addSignal = (code: SignalCode, weight: number, evidence: string) => {
    if (seen.has(code)) return
    seen.add(code)
    signals.push({ code, weight, evidence })
  }

  for (const rule of textRules) {
    const match = input.match(rule.pattern)
    if (match) addSignal(rule.code, rule.weight, match[0])
  }

  for (const rawUrl of urls) {
    const parsed = asUrl(rawUrl)
    if (!parsed) {
      addSignal('obfuscatedUrl', 18, rawUrl)
      continue
    }

    const hostname = parsed.hostname.toLowerCase().replace(/^www\./, '')
    const tld = hostname.split('.').at(-1) ?? ''

    if (suspiciousTlds.has(tld)) addSignal('suspiciousTld', 19, `.${tld}`)
    if (shorteners.has(hostname)) addSignal('shortener', 15, hostname)
    if (/^(?:\d{1,3}\.){3}\d{1,3}$/.test(hostname)) addSignal('ipAddress', 26, hostname)
    if (hostname.includes('xn--')) addSignal('punycode', 22, hostname)
    if (rawUrl.toLowerCase().startsWith('http://')) addSignal('insecureProtocol', 10, 'http://')
    if (rawUrl.includes('@') || (hostname.match(/-/g)?.length ?? 0) >= 3 || rawUrl.length > 120) {
      addSignal('obfuscatedUrl', 15, hostname)
    }

    for (const brand of brands) {
      if (!brand.pattern.test(input)) continue
      const isOfficial = brand.domains.some(
        (domain) => hostname === domain || hostname.endsWith(`.${domain}`),
      )
      if (!isOfficial) addSignal('brandMismatch', 30, hostname)
    }
  }

  let score = signals.reduce((total, signal) => total + signal.weight, 0)
  if (signals.length >= 3) score += 9
  if (signals.length >= 5) score += 7
  score = Math.min(99, Math.max(signals.length ? 12 : 4, score))

  return {
    id: createId(),
    input,
    mode,
    score,
    level: levelFromScore(score),
    signals: signals.sort((a, b) => b.weight - a.weight),
    actions: actionList(signals),
    urls,
    createdAt: new Date().toISOString(),
  }
}
