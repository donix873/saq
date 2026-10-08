export type Locale = 'ru' | 'kk'

export type ScanMode = 'message' | 'link'

export type RiskLevel = 'low' | 'medium' | 'high' | 'critical'

export type SignalCode =
  | 'urgency'
  | 'threat'
  | 'impersonation'
  | 'secretRequest'
  | 'personalData'
  | 'paymentRequest'
  | 'tooGoodToBeTrue'
  | 'suspiciousJob'
  | 'isolation'
  | 'remoteAccess'
  | 'suspiciousTld'
  | 'shortener'
  | 'ipAddress'
  | 'punycode'
  | 'insecureProtocol'
  | 'obfuscatedUrl'
  | 'brandMismatch'

export type ActionCode =
  | 'dontOpenLink'
  | 'dontShareCodes'
  | 'dontShareData'
  | 'stopPayment'
  | 'dontInstallApps'
  | 'verifyOfficially'
  | 'contactSupport'
  | 'preserveEvidence'
  | 'verifySender'

export interface RiskSignal {
  code: SignalCode
  weight: number
  evidence: string
}

export interface AnalysisResult {
  id: string
  input: string
  mode: ScanMode
  score: number
  level: RiskLevel
  signals: RiskSignal[]
  actions: ActionCode[]
  urls: string[]
  createdAt: string
}
