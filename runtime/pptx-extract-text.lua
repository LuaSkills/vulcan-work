-- Read one office document selection through the fixed managed Python handler.
-- 通过固定的受管 Python 处理器读取一段办公文档。
-- @param args table Local file path and optional selection defined by the entry schema.
-- @param args table 入口 Schema 定义的本地文件路径和可选范围。
-- @return string Structured result encoded as JSON.
-- @return string 编码为 JSON 的结构化结果。
return function(args)
  -- Run the package-owned handler with the host's execution deadline.
  -- 在宿主执行时限内调用技能包自有处理器。
  local result = vulcan.runtime.python.invoke({
    file = "python/office_documents.py",
    handler = "pptx_extract_text",
    args = args,
    timeout_ms = 60000,
  })
  if not result.ok then
    error(tostring(result.error))
  end
  return vulcan.json.encode(result.value)
end
