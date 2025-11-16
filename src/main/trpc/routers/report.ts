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
  OutputWriteError,
  MissingSourceError
} from '../../services/errors'
import { createLogger } from '../../services/logger'

const log = createLogger('reportRouter')

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
  /** 模板额外数据源（可选，key -> 文件路径） */
  extraSources: z.record(z.string(), z.string()).optional(),
  /** Carbone 渲染选项（可选） */
  renderOptions: z.record(z.string(), z.unknown()).optional(),
  /** 用户输入参数（可选，如年月等模板特定参数） */
  userInput: z.record(z.string(), z.unknown()).optional()
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
    log.info('开始生成报表', {
      templateId: input.templateId,
      sourcePath: input.sourcePath,
      outputDir: input.outputDir
    })

    try {
      // 步骤 1: Excel → 数据
      const parseResult = await excelToData({
        sourcePath: input.sourcePath,
        templateId: input.templateId,
        parseOptions: input.parseOptions,
        extraSources: input.extraSources
      })

      log.info('Excel 解析完成', {
        templateId: input.templateId,
        warningCount: parseResult.warnings.length,
        sheets: parseResult.sourceMeta.sheets
      })

      // 步骤 2: 数据 → 报表
      const reportResult = await dataToReport({
        templateId: input.templateId,
        parsedData: parseResult.data,
        outputDir: input.outputDir,
        reportName: input.reportName,
        renderOptions: input.renderOptions,
        userInput: input.userInput
      })

      const duration = ((Date.now() - startTime) / 1000).toFixed(2)
      log.info('报表生成成功', {
        templateId: input.templateId,
        durationSeconds: Number(duration),
        outputPath: reportResult.outputPath,
        size: reportResult.size
      })

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
      log.error('报表生成失败', {
        templateId: input.templateId,
        error
      })

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

      if (error instanceof MissingSourceError) {
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
