-- Invoke the managed document reader with arguments validated by the entry schema.
-- 使用入口 Schema 已校验的参数调用受管文档读取器。
-- @param args table Input values declared by the entry schema.
-- @param args table 入口 Schema 声明的输入值。
-- @return string Structured tool output encoded as JSON.
-- @return string 编码为 JSON 的结构化工具结果。
return function(args)
  -- The fixed Python handler only exposes the operation named by this entry.
  -- 固定的 Python 处理器只暴露当前入口对应的操作。
  local result = vulcan.runtime.python.invoke({
    file = "python/office.py",
    handler = "xlsx_read_range",
    args = args,
    timeout_ms = 60000,
  })
  if not result.ok then
    error(tostring(result.error))
  end
  return vulcan.json.encode(result.value)
end
