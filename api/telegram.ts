import { timingSafeEqual } from 'node:crypto'
import type { IncomingMessage, ServerResponse } from 'node:http'

const DEFAULT_APP_URL = 'https://saq-nine.vercel.app'

type TelegramChat = {
  id: number
}

type TelegramMessage = {
  chat?: TelegramChat
  text?: string
}

export type TelegramUpdate = {
  message?: TelegramMessage
}

type TelegramRequest = IncomingMessage & {
  body?: TelegramUpdate | string
}

export type BotReply = {
  chat_id: number
  text: string
  parse_mode: 'HTML'
  disable_web_page_preview: boolean
  reply_markup: {
    inline_keyboard: Array<Array<{ text: string; url: string }>>
  }
}

const welcomeText = `🛡 <b>SAQ — цифрлық қорғаныс</b>

Сәлем! SAQ күмәнді хабарламалар мен сілтемелерді тексеруге көмектеседі.

Здравствуйте! SAQ помогает проверить подозрительные сообщения и ссылки на признаки мошенничества.

Төмендегі батырманы басыңыз / Нажмите кнопку ниже 👇`

export function createBotReply(update: TelegramUpdate, appUrl = DEFAULT_APP_URL): BotReply | null {
  const chatId = update.message?.chat?.id
  const command = update.message?.text?.trim().split(/\s+/)[0]?.split('@')[0].toLowerCase()

  if (!chatId || (command !== '/start' && command !== '/help')) {
    return null
  }

  return {
    chat_id: chatId,
    text: welcomeText,
    parse_mode: 'HTML',
    disable_web_page_preview: false,
    reply_markup: {
      inline_keyboard: [
        [
          {
            text: '🛡 SAQ-ты ашу / Открыть SAQ',
            url: appUrl,
          },
        ],
      ],
    },
  }
}

function sendJson(response: ServerResponse, status: number, body: unknown) {
  response.statusCode = status
  response.setHeader('content-type', 'application/json; charset=utf-8')
  response.end(JSON.stringify(body))
}

function hasValidSecret(request: TelegramRequest, expectedSecret: string) {
  const header = request.headers['x-telegram-bot-api-secret-token']
  const receivedSecret = Array.isArray(header) ? header[0] : header

  if (!receivedSecret) return false

  const received = Buffer.from(receivedSecret)
  const expected = Buffer.from(expectedSecret)

  return received.length === expected.length && timingSafeEqual(received, expected)
}

async function readUpdate(request: TelegramRequest): Promise<TelegramUpdate> {
  if (request.body && typeof request.body !== 'string') return request.body
  if (typeof request.body === 'string') return JSON.parse(request.body) as TelegramUpdate

  const chunks: Buffer[] = []
  for await (const chunk of request) chunks.push(Buffer.from(chunk))

  const body = Buffer.concat(chunks).toString('utf8')
  return body ? (JSON.parse(body) as TelegramUpdate) : {}
}

async function callTelegram(token: string, method: string, payload: unknown) {
  const result = await fetch(`https://api.telegram.org/bot${token}/${method}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(payload),
  })

  const data = (await result.json()) as { ok?: boolean; description?: string }
  if (!result.ok || !data.ok) {
    throw new Error(data.description || `Telegram API returned ${result.status}`)
  }

  return data
}

function normalizeAppUrl(value: string | undefined) {
  return (value || DEFAULT_APP_URL).replace(/\/$/, '')
}

export default async function handler(request: TelegramRequest, response: ServerResponse) {
  const token = process.env.TELEGRAM_BOT_TOKEN
  const webhookSecret = process.env.TELEGRAM_WEBHOOK_SECRET
  const appUrl = normalizeAppUrl(process.env.SAQ_APP_URL)

  if (!token || !webhookSecret) {
    sendJson(response, 503, {
      ok: false,
      error: 'Telegram bot is not configured in Vercel environment variables.',
    })
    return
  }

  if (!/^[A-Za-z0-9_-]{16,256}$/.test(webhookSecret)) {
    sendJson(response, 503, {
      ok: false,
      error: 'TELEGRAM_WEBHOOK_SECRET must contain 16-256 letters, numbers, underscores or hyphens.',
    })
    return
  }

  try {
    if (request.method === 'GET') {
      await callTelegram(token, 'setWebhook', {
        url: `${appUrl}/api/telegram`,
        secret_token: webhookSecret,
        allowed_updates: ['message'],
        drop_pending_updates: false,
      })
      await callTelegram(token, 'setMyCommands', {
        commands: [
          { command: 'start', description: 'SAQ қолданбасын ашу / Открыть SAQ' },
          { command: 'help', description: 'Көмек / Помощь' },
        ],
      })

      sendJson(response, 200, {
        ok: true,
        message: 'Webhook configured. Open @saq_kz_bot and send /start.',
      })
      return
    }

    if (request.method !== 'POST') {
      response.setHeader('allow', 'GET, POST')
      sendJson(response, 405, { ok: false, error: 'Method not allowed.' })
      return
    }

    if (!hasValidSecret(request, webhookSecret)) {
      sendJson(response, 401, { ok: false, error: 'Invalid webhook secret.' })
      return
    }

    const update = await readUpdate(request)
    const reply = createBotReply(update, appUrl)

    if (reply) await callTelegram(token, 'sendMessage', reply)

    sendJson(response, 200, { ok: true })
  } catch (error) {
    console.error('Telegram webhook error:', error)
    sendJson(response, 500, { ok: false, error: 'Telegram request failed.' })
  }
}
