# Known math limitations

测试已知不支持结构的哨点。

以下结构预期转换成功，但 OneNote 渲染异常（换行无效），对应 `docs/design/technical_debt/latex-mathml.md` 中的未解决问题。若行为变好或变坏，应同步更新该文档。E2E 回归时此文件不计入“全绿”门槛。

## 非矩阵环境换行（预期换行无效）

$$a \\ b$$
