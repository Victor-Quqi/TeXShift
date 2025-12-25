# OneNote 错误代码参考

本文档列出了 OneNote 2013 对象模型中的所有错误代码及其说明。

> 来源：https://learn.microsoft.com/zh-cn/office/client-developer/onenote/error-codes-onenote

## 基础错误代码

| HRESULT | 值 | 说明 |
|---------|-----|------|
| hrMalformedXML | 0x80042000 | XML 格式不正确 |
| hrInvalidXML | 0x80042001 | XML 无效 |
| hrCreatingSection | 0x80042002 | 无法创建部分 |
| hrOpeningSection | 0x80042003 | 不能打开部分 |
| hrSectionDoesNotExist | 0x80042004 | 节不存在 |
| hrPageDoesNotExist | 0x80042005 | 页面不存在 |
| hrFileDoesNotExist | 0x80042006 | 文件不存在 |
| hrInsertingImage | 0x80042007 | 无法插入图像 |
| hrInsertingInk | 0x80042008 | 无法插入墨迹 |
| hrInsertingHtml | 0x80042009 | 无法插入 HTML |
| hrNavigationToPage | 0x8004200a | 无法打开页面 |
| hrSectionReadOnly | 0x8004200b | 部分是只读的 |
| hrPageReadOnly | 0x8004200c | 页是只读的 |
| hrInsertingOutlineText | 0x8004200d | 无法插入大纲文本 |
| hrPageObjectDoesNotExist | 0x8004200e | Page 对象不存在 |
| hrBinaryObjectDoesNotExist | 0x8004200f | 二进制对象不存在 |
| hrLastModifiedDateDidNotMatch | 0x80042010 | 上次修改的日期不符 |
| hrGroupDoesNotExist | 0x80042011 | 节组不存在 |
| hrPageDoesNotExistInGroup | 0x80042012 | 页面不存在该节组中 |
| hrNoActiveSelection | 0x80042013 | 没有任何活动的选定内容 |
| hrObjectDoesNotExist | 0x80042014 | 对象不存在 |
| hrNotebookDoesNotExist | 0x80042015 | 笔记本不存在 |
| hrInsertingFile | 0x80042016 | 无法插入该文件 |
| hrInvalidName | 0x80042017 | 名称无效 |
| hrFolderDoesNotExist | 0x80042018 | 文件夹（部分组）不存在 |
| hrInvalidQuery | 0x80042019 | 查询无效 |
| hrFileAlreadyExists | 0x8004201a | 文件已存在 |
| hrSectionEncryptedAndLocked | 0x8004201b | 部分已加密并锁定 |
| hrDisabledByPolicy | 0x8004201c | 操作已被策略禁用 |

## 扩展错误代码

| HRESULT | 值 | 说明 |
|---------|-----|------|
| hrNotYetSynchronized | 0x8004201d | OneNote 尚未同步内容 |
| hrLegacySection | 0x8004201e | 部分来自 OneNote 2007 或更早版本 |
| hrMergeFailed | 0x8004201f | 合并操作失败 |
| hrInvalidXMLSchema | 0x80042020 | XML 架构无效 |
| hrFutureContentLoss | 0x80042022 | 内容丢失（来自 OneNote 的未来版本） |
| hrTimeOut | 0x80042023 | 操作超时 |
| hrRecordingInProgress | 0x80042024 | 音频录制正在进行 |
| hrUnknownLinkedNoteState | 0x80042025 | 链接笔记状态未知 |
| hrNoShortNameForLinkedNote | 0x80042026 | 链接笔记没有短名称 |
| hrNoFriendlyNameForLinkedNote | 0x80042027 | 链接笔记没有友好名称 |
| hrInvalidLinkedNoteUri | 0x80042028 | 链接笔记 URI 无效 |
| hrInvalidLinkedNoteThumbnail | 0x80042029 | 链接笔记缩略图无效 |
| hrImportLNTThumbnailFailed | 0x8004202a | 链接笔记缩略图导入失败 |
| hrUnreadDisabledForNotebook | 0x8004202b | 笔记本的未读突出显示已禁用 |
| hrInvalidSelection | 0x8004202c | 所选内容无效 |
| hrConvertFailed | 0x8004202d | 转换失败 |
| hrRecycleBinEditFailed | 0x8004202e | 在回收站中编辑失败 |

## OneNote 2013 新增错误代码

| HRESULT | 值 | 说明 |
|---------|-----|------|
| hrIMConversationTypeInvalid | 0x8004202f | IMConversationType 页面节点属性值无效（需为 0、1、2 或 3） |
| hrAppInModalUI | 0x80042030 | 模式对话框阻止应用程序 |
| hrPublishFormatUnsupportedForLabels | 0x80042031 | 发布格式不支持敏感度标记 |
