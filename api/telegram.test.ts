import { describe, expect, it } from 'vitest'
import { createBotReply } from './telegram'

describe('Telegram bot reply', () => {
  it('creates a bilingual website button for /start', () => {
    const reply = createBotReply(
      { message: { chat: { id: 873 }, text: '/start' } },
      'https://saq-nine.vercel.app',
    )

    expect(reply?.chat_id).toBe(873)
    expect(reply?.text).toContain('SAQ — цифрлық қорғаныс')
    expect(reply?.text).toContain('Здравствуйте!')
    expect(reply?.reply_markup.inline_keyboard[0][0]).toEqual({
      text: '🛡 SAQ-ты ашу / Открыть SAQ',
      url: 'https://saq-nine.vercel.app',
    })
  })

  it('accepts a command addressed to the bot', () => {
    const reply = createBotReply({
      message: { chat: { id: 1 }, text: '/start@saq_kz_bot referral' },
    })

    expect(reply).not.toBeNull()
  })

  it('ignores unrelated messages', () => {
    const reply = createBotReply({ message: { chat: { id: 1 }, text: 'Сәлем' } })
    expect(reply).toBeNull()
  })
})
