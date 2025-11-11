/**
 * 模板注册中心
 * 负责模板的注册、获取、列举与校验
 */

import type { TemplateId, TemplateDefinition, TemplateMeta } from './types'
import { TemplateNotFoundError } from '../errors'
import { templateExists } from '../utils/filePaths'

/**
 * 全局模板注册表
 */
const templateRegistry = new Map<TemplateId, TemplateDefinition>()

/**
 * 注册模板
 * @param definition 模板完整定义
 */
export function registerTemplate(definition: TemplateDefinition): void {
  const { id } = definition.meta
  if (templateRegistry.has(id)) {
    console.warn(`[Registry] 模板 ${id} 已存在，将被覆盖`)
  }
  templateRegistry.set(id, definition)
  console.log(`[Registry] 模板已注册: ${id} - ${definition.meta.name}`)
}

/**
 * 获取模板定义
 * @param id 模板 ID
 * @throws TemplateNotFoundError 模板不存在时抛出
 */
export function getTemplate(id: TemplateId): TemplateDefinition {
  const definition = templateRegistry.get(id)
  if (!definition) {
    throw new TemplateNotFoundError(id)
  }
  return definition
}

/**
 * 列出所有已注册的模板元信息
 */
export function listTemplates(): TemplateMeta[] {
  console.log(`[Registry] 列出所有已注册的模板元信息`, Array.from(templateRegistry.values()))
  return Array.from(templateRegistry.values()).map((def) => def.meta)
}

/**
 * 校验模板文件是否存在
 * @param id 模板 ID
 * @returns 校验结果 { ok: boolean, issues?: string[] }
 */
export function validateTemplate(id: TemplateId): { ok: boolean; issues?: string[] } {
  const definition = templateRegistry.get(id)
  if (!definition) {
    return { ok: false, issues: [`模板未注册: ${id}`] }
  }

  const issues: string[] = []

  // 检查模板文件是否存在
  if (!templateExists(definition.meta.filename)) {
    issues.push(`模板文件不存在: ${definition.meta.filename}`)
  }

  return {
    ok: issues.length === 0,
    issues: issues.length > 0 ? issues : undefined
  }
}

/**
 * 清空注册表（测试用）
 */
export function clearRegistry(): void {
  templateRegistry.clear()
}
