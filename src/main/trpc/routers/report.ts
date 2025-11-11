/**
 * Report Router
 * 提供报表生成相关的 tRPC 接口
 */

import { z } from 'zod'
import { router, publicProcedure, TRPCError } from '../trpc'
import { excelToData } from '../../services/excelToData'
import { dataToReport } from '../../services/dataToReport'
import {
  AppError,
  TemplateNotFoundError,
  UnsupportedFileError,
  ExcelFileTooLargeError,
  ExcelParseError,
  ReportRenderError,
  OutputDirNotSelectedError,
  OutputWriteError
} from '../../services/errors'

/**
 * 报表生成输入参数 Schema
 */
const generateInputSchema = z.object({
  /** 模板 ID */
  templateId: z.string().min(1, '模板 ID 不能为空'),
  /** 源文件路径 */
  sourcePath: z.string().min(1, '源文件路径不能为空'),
  /** 输出目录 */
  outputDir: z.string().min(1, '输出目录不能为空'),
  /** 报表名称（可选，若无则使用兜底策略） */
  reportName: z.string().optional(),
  /** 模板解析选项（可选） */
  parseOptions: z.record(z.string(), z.unknown()).optional(),
  /** Carbone 渲染选项（可选） */
  renderOptions: z.record(z.string(), z.unknown()).optional()
})

/**
 * Report Router
 */
export const reportRouter = router({
  /**
   * 生成报表（同步返回结果）
   */
  generate: publicProcedure.input(generateInputSchema).mutation(async ({ input }) => {
    const startTime = Date.now()
    console.log('[Report Router] 开始生成报表:', input.templateId)

    try {
      // 步骤 1: Excel → 数据
      const parseResult = await excelToData({
        sourcePath: input.sourcePath,
        templateId: input.templateId,
        parseOptions: input.parseOptions
      })

      console.log(`[Report Router] 数据解析完成，警告数量: ${parseResult.warnings.length}`)

      // 步骤 2: 数据 → 报表
      const reportResult = await dataToReport({
        templateId: input.templateId,
        parsedData: parseResult.data,
        outputDir: input.outputDir,
        reportName: input.reportName,
        renderOptions: input.renderOptions
      })

      const duration = Date.now() - startTime
      console.log(
        `[Report Router] 报表生成成功，耗时: ${duration}ms，输出: ${reportResult.outputPath}`
      )

      // 返回成功结果
      return {
        success: true,
        outputPath: reportResult.outputPath,
        size: reportResult.size,
        generatedAt: reportResult.generatedAt,
        warnings: parseResult.warnings,
        duration
      }
    } catch (error) {
      console.error('[Report Router] 报表生成失败:', error)

      // 映射自定义错误到 TRPCError
      if (error instanceof TemplateNotFoundError) {
        throw new TRPCError({
          code: 'NOT_FOUND',
          message: error.message,
          cause: error
        })
      }

      if (error instanceof UnsupportedFileError) {
        throw new TRPCError({
          code: 'BAD_REQUEST',
          message: error.message,
          cause: error
        })
      }

      if (error instanceof ExcelFileTooLargeError) {
        throw new TRPCError({
          code: 'BAD_REQUEST',
          message: error.message,
          cause: error
        })
      }

      if (error instanceof ExcelParseError) {
        throw new TRPCError({
          code: 'BAD_REQUEST',
          message: error.message,
          cause: error
        })
      }

      if (error instanceof ReportRenderError) {
        throw new TRPCError({
          code: 'INTERNAL_SERVER_ERROR',
          message: error.message,
          cause: error
        })
      }

      if (error instanceof OutputDirNotSelectedError) {
        throw new TRPCError({
          code: 'BAD_REQUEST',
          message: error.message,
          cause: error
        })
      }

      if (error instanceof OutputWriteError) {
        throw new TRPCError({
          code: 'INTERNAL_SERVER_ERROR',
          message: error.message,
          cause: error
        })
      }

      // 兜底错误处理
      if (error instanceof AppError) {
        throw new TRPCError({
          code: 'INTERNAL_SERVER_ERROR',
          message: error.message,
          cause: error
        })
      }

      // 未知错误
      throw new TRPCError({
        code: 'INTERNAL_SERVER_ERROR',
        message: error instanceof Error ? error.message : '报表生成失败',
        cause: error
      })
    }
  })
})
