import type { AnalysisResult } from '../types'

const STORAGE_KEY = 'saq.scan.history.v1'
const HISTORY_LIMIT = 20

function isAnalysisResult(value: unknown): value is AnalysisResult {
  if (!value || typeof value !== 'object') return false
  const candidate = value as Partial<AnalysisResult>

  return (
    typeof candidate.id === 'string' &&
    typeof candidate.input === 'string' &&
    typeof candidate.score === 'number' &&
    typeof candidate.createdAt === 'string' &&
    Array.isArray(candidate.signals) &&
    Array.isArray(candidate.actions)
  )
}

export function loadHistory(): AnalysisResult[] {
  if (typeof window === 'undefined') return []

  try {
    const stored = window.localStorage.getItem(STORAGE_KEY)
    if (!stored) return []
    const parsed: unknown = JSON.parse(stored)
    return Array.isArray(parsed) ? parsed.filter(isAnalysisResult).slice(0, HISTORY_LIMIT) : []
  } catch {
    return []
  }
}

export function saveHistory(history: AnalysisResult[]): boolean {
  if (typeof window === 'undefined') return false

  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(history.slice(0, HISTORY_LIMIT)))
    return true
  } catch {
    return false
  }
}

export function clearHistory(): boolean {
  if (typeof window === 'undefined') return false

  try {
    window.localStorage.removeItem(STORAGE_KEY)
    return true
  } catch {
    return false
  }
}
