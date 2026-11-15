-- Dispatch the office operation to its fixed package-owned handler.
-- 将办公操作交给技能包中固定的处理器。
-- @param args table Arguments declared by the entry schema.
-- @param args table 入口 Schema 声明的参数。
-- @return string JSON result from the managed runtime.
-- @return string 受管运行时返回的 JSON 结果。
return function(args)
  -- Surface a failed managed runtime operation through the tool error.
  -- 将受管运行时操作失败作为工具错误返回。
  local result = vulcan.runtime.python.invoke({
    file = "python/pptx_structure.py",
    handler = "pptx_remove_slides",
    args = args,
    timeout_ms = 60000,
  })
  if not result.ok then
    error(tostring(result.error))
  end
  return vulcan.json.encode(result.value)
end
