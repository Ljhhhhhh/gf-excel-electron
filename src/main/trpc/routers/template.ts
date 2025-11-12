/**
 * Template Router
 * 提供模板管理相关的 tRPC 接口
 */

import { z } from 'zod'
import { router, publicProcedure, TRPCError } from '../trpc'
import { listTemplates, validateTemplate, getTemplate } from '../../services/templates/registry'
import { TemplateNotFoundError } from '../../services/errors'

/**
 * Template Router
 */
export const templateRouter = router({
  /**
   * 列出所有已注册的模板
   */
  list: publicProcedure.query(() => {
    try {
      return listTemplates()
    } catch (error) {
      console.error('[Template Router] list failed:', error)
      throw new TRPCError({
        code: 'INTERNAL_SERVER_ERROR',
        message: '获取模板列表失败'
      })
    }
  }),

  /**
   * 获取模板元信息
   */
  getMeta: publicProcedure
    .input(
      z.object({
        id: z.string().min(1, '模板 ID 不能为空')
      })
    )
    .query(({ input }) => {
      try {
        const template = getTemplate(input.id)
        return template.meta
      } catch (error) {
        if (error instanceof TemplateNotFoundError) {
          throw new TRPCError({
            code: 'NOT_FOUND',
            message: error.message
          })
        }
        console.error('[Template Router] getMeta failed:', error)
        throw new TRPCError({
          code: 'INTERNAL_SERVER_ERROR',
          message: '获取模板信息失败'
        })
      }
    }),

  /**
   * 校验模板文件是否存在
   */
  validate: publicProcedure
    .input(
      z.object({
        id: z.string().min(1, '模板 ID 不能为空')
      })
    )
    .query(({ input }) => {
      try {
        return validateTemplate(input.id)
      } catch (error) {
        console.error('[Template Router] validate failed:', error)
        throw new TRPCError({
          code: 'INTERNAL_SERVER_ERROR',
          message: '校验模板失败'
        })
      }
    }),

  /**
   * 获取模板的 inputRule（用于前端渲染表单）
   */
  getInputRule: publicProcedure
    .input(
      z.object({
        templateId: z.string().min(1, '模板 ID 不能为空')
      })
    )
    .query(({ input }) => {
      try {
        const template = getTemplate(input.templateId)

        if (!template) {
          throw new TemplateNotFoundError(input.templateId)
        }

        // 返回 formCreate rule（可序列化的 JSON）
        return {
          templateId: input.templateId,
          inputRule: template.inputRule || null
        }
      } catch (error) {
        if (error instanceof TemplateNotFoundError) {
          throw new TRPCError({
            code: 'NOT_FOUND',
            message: error.message
          })
        }
        console.error('[Template Router] getInputRule failed:', error)
        throw new TRPCError({
          code: 'INTERNAL_SERVER_ERROR',
          message: '获取模板输入规则失败'
        })
      }
    })
})
