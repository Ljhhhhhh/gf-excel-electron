/**
 * File Router
 * 提供文件选择和文件操作相关的 tRPC 接口
 */

import { z } from 'zod'
import { dialog } from 'electron'
import { router, publicProcedure, TRPCError } from '../trpc'
import { openFolder } from '../../services/utils/fileOps'
import fs from 'node:fs'

/**
 * File Router
 */
export const fileRouter = router({
  /**
   * 打开文件选择对话框，选择源 Excel 文件
   */
  selectSourceFile: publicProcedure.query(async () => {
    try {
      const result = await dialog.showOpenDialog({
        title: '选择源数据文件',
        properties: ['openFile'],
        filters: [
          { name: 'Excel 文件', extensions: ['xlsx', 'xls'] },
          { name: '所有文件', extensions: ['*'] }
        ]
      })

      if (result.canceled || result.filePaths.length === 0) {
        return { canceled: true, filePath: null }
      }

      return {
        canceled: false,
        filePath: result.filePaths[0]
      }
    } catch (error) {
      console.error('[File Router] selectSourceFile failed:', error)
      throw new TRPCError({
        code: 'INTERNAL_SERVER_ERROR',
        message: '打开文件选择对话框失败'
      })
    }
  }),

  /**
   * 打开目录选择对话框，选择输出目录
   */
  selectOutputDir: publicProcedure.query(async () => {
    try {
      const result = await dialog.showOpenDialog({
        title: '选择输出目录',
        properties: ['openDirectory', 'createDirectory']
      })

      if (result.canceled || result.filePaths.length === 0) {
        return { canceled: true, dirPath: null }
      }

      return {
        canceled: false,
        dirPath: result.filePaths[0]
      }
    } catch (error) {
      console.error('[File Router] selectOutputDir failed:', error)
      throw new TRPCError({
        code: 'INTERNAL_SERVER_ERROR',
        message: '打开目录选择对话框失败'
      })
    }
  }),

  /**
   * 在文件管理器中打开指定路径
   */
  openInFolder: publicProcedure
    .input(
      z.object({
        path: z.string().min(1, '路径不能为空')
      })
    )
    .mutation(async ({ input }) => {
      try {
        // 校验路径是否存在
        if (!fs.existsSync(input.path)) {
          throw new TRPCError({
            code: 'NOT_FOUND',
            message: '文件或文件夹不存在'
          })
        }

        await openFolder(input.path)
        return { success: true }
      } catch (error) {
        if (error instanceof TRPCError) {
          throw error
        }
        console.error('[File Router] openInFolder failed:', error)
        throw new TRPCError({
          code: 'INTERNAL_SERVER_ERROR',
          message: '打开文件夹失败'
        })
      }
    })
})
