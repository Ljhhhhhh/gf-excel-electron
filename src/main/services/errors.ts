/**
 * 统一错误定义
 * 所有自定义错误继承自 AppError，携带 code/message/details
 */

export class AppError extends Error {
  constructor(
    public code: string,
    message: string,
    public details?: unknown
  ) {
    super(message)
    this.name = this.constructor.name
    Error.captureStackTrace(this, this.constructor)
  }
}

// ========== 模板相关错误 ==========

export class TemplateNotFoundError extends AppError {
  constructor(templateId: string) {
    super('TEMPLATE_NOT_FOUND', `模板未找到: ${templateId}`, { templateId })
  }
}

export class TemplateUnsupportedExtError extends AppError {
  constructor(ext: string, allowedExts: string[]) {
    super('TEMPLATE_UNSUPPORTED_EXT', `不支持的模板扩展名: ${ext}`, {
      ext,
      allowedExts
    })
  }
}

// ========== Excel 文件相关错误 ==========

export class UnsupportedFileError extends AppError {
  constructor(filePath: string, reason: string) {
    super('EXCEL_UNSUPPORTED_FILE', `不支持的文件: ${filePath}，原因: ${reason}`, {
      filePath,
      reason
    })
  }
}

export class ExcelFileTooLargeError extends AppError {
  constructor(filePath: string, size: number, limit: number) {
    super(
      'EXCEL_FILE_TOO_LARGE',
      `Excel 文件过大: ${filePath}，大小: ${(size / 1024 / 1024).toFixed(2)}MB，限制: ${limit}MB`,
      { filePath, size, limit }
    )
  }
}

export class ExcelParseError extends AppError {
  constructor(filePath: string, originalError: unknown) {
    super('EXCEL_PARSE_ERROR', `Excel 解析失败: ${filePath}`, {
      filePath,
      originalError: originalError instanceof Error ? originalError.message : String(originalError)
    })
  }
}

// ========== 报表渲染相关错误 ==========

export class ReportRenderError extends AppError {
  constructor(templateId: string, originalError: unknown) {
    super('REPORT_RENDER_ERROR', `报表渲染失败: ${templateId}`, {
      templateId,
      originalError: originalError instanceof Error ? originalError.message : String(originalError)
    })
  }
}

// ========== 输出相关错误 ==========

export class OutputDirNotSelectedError extends AppError {
  constructor() {
    super('OUTPUT_DIR_NOT_SELECTED', '输出目录未选择')
  }
}

export class OutputWriteError extends AppError {
  constructor(outputPath: string, originalError: unknown) {
    super('OUTPUT_WRITE_ERROR', `写入输出文件失败: ${outputPath}`, {
      outputPath,
      originalError: originalError instanceof Error ? originalError.message : String(originalError)
    })
  }
}

// ========== Warning 结构 ==========

export interface Warning {
  code: string
  message: string
  level?: 'info' | 'warn'
  context?: unknown
}
