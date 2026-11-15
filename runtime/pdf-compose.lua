-- Produce a new office artifact through the package's fixed write handler.
-- 通过技能包中固定的写入处理器生成新的办公文件。
-- @param args table Inputs and explicit output path declared in the entry schema.
-- @param args table 入口 Schema 声明的输入和显式输出路径。
-- @return string JSON metadata for the published file.
-- @return string 已发布文件的 JSON 元数据。
return function(args)
  -- Errors are surfaced so users can correct a failed write without replacing any source.
  -- 显式返回失败，让用户能纠正错误并保留源文件。
  local result = vulcan.runtime.python.invoke({
    file = "python/write_pdf_xlsx.py",
    handler = "pdf_compose",
    args = args,
    timeout_ms = 60000,
  })
  if not result.ok then
    error(tostring(result.error))
  end
  return vulcan.json.encode(result.value)
end
