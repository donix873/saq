import { describe, expect, it } from 'vitest'
import { analyzeInput } from './analyzer'

describe('analyzeInput', () => {
  it('marks a bank impersonation message with a fake domain as critical', () => {
    const result = analyzeInput(
      'Служба безопасности Kaspi: срочно подтвердите данные по ссылке http://kaspi-secure-pay.cc/auth, иначе счёт будет заблокирован. Никому не говорите.',
      'message',
    )

    expect(result.level).toBe('critical')
    expect(result.score).toBeGreaterThanOrEqual(75)
    expect(result.signals.map((signal) => signal.code)).toEqual(
      expect.arrayContaining(['urgency', 'threat', 'impersonation', 'brandMismatch']),
    )
    expect(result.actions).toContain('dontOpenLink')
    expect(result.actions).toContain('verifyOfficially')
  })

  it('keeps a normal planning message at low risk', () => {
    const result = analyzeInput(
      'Привет! Встреча завтра в 15:00 в школьной библиотеке. Если планы изменятся, напиши.',
      'message',
    )

    expect(result.level).toBe('low')
    expect(result.signals).toHaveLength(0)
    expect(result.actions).toEqual(['verifySender'])
  })

  it('recognizes secret-code requests in Kazakh', () => {
    const result = analyzeInput(
      'Шотыңызды қорғау үшін қазір SMS растау кодын және карта деректерін жіберіңіз.',
      'message',
    )

    expect(result.signals.map((signal) => signal.code)).toEqual(
      expect.arrayContaining(['secretRequest', 'personalData', 'urgency']),
    )
    expect(result.actions).toContain('dontShareCodes')
    expect(result.actions).toContain('dontShareData')
  })

  it('detects a raw IP and insecure protocol', () => {
    const result = analyzeInput('http://192.168.10.10/login', 'link')

    expect(result.signals.map((signal) => signal.code)).toEqual(
      expect.arrayContaining(['ipAddress', 'insecureProtocol']),
    )
    expect(result.actions).toContain('dontOpenLink')
  })

  it('does not flag the official Kaspi domain as a mismatch', () => {
    const result = analyzeInput('Kaspi: https://kaspi.kz', 'link')

    expect(result.signals.map((signal) => signal.code)).not.toContain('brandMismatch')
  })

  it('detects remote-access software requests', () => {
    const result = analyzeInput(
      'Установите AnyDesk и покажите экран. Никому не говорите и не кладите трубку.',
      'message',
    )

    expect(result.signals.map((signal) => signal.code)).toEqual(
      expect.arrayContaining(['remoteAccess', 'isolation']),
    )
    expect(result.actions).toContain('dontInstallApps')
  })

  it('deduplicates repeated signals', () => {
    const result = analyzeInput('Срочно, срочно, срочно подтвердите данные прямо сейчас.', 'message')
    const urgencySignals = result.signals.filter((signal) => signal.code === 'urgency')

    expect(urgencySignals).toHaveLength(1)
  })

  it('rejects input that is too short to analyze', () => {
    expect(() => analyzeInput('ok', 'message')).toThrow('INPUT_TOO_SHORT')
  })
})
