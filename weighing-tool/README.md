# 称重单识别整理工具

用于将称重单图片或夸克识别后的 Excel 整理成固定栏位工作簿。

## 功能

- 支持拖入图片、Excel 或整个文件夹。
- 每张识别结果生成一个标签页。
- 左侧保留夸克返回原表，右侧生成整理区。
- 右侧整理区通过公式引用左侧原表，左侧修改后右侧会跟随变化。
- 栏位固定为 `1栏南` 到 `5栏北`，以及 `公鸡1` 到 `公鸡5`。
- 每列底部显示识别数量和核对结果。
- 可手动选择栋舍：`1-1`、`1-2`、`2-1`、`2-2`、`3-1`、`3-2`、`4`、`5`、`6`。

## 夸克 API Key 配置

不要把真实 Key 提交到 GitHub。

复制 `quark_credentials.example.json` 为 `quark_credentials.json`，填写：

```json
{
  "client_id": "你的 API Key ID",
  "client_secret": "你的 API Key"
}
```

也可以用环境变量：

- `QUARK_SCAN_CLIENT_ID`
- `QUARK_SCAN_CLIENT_SECRET`

## 运行

开发运行：

```powershell
python 称重单整理工具_重做版.py
```

打包后运行：

- 将 `quark_credentials.json` 放在 exe 同目录。
- 双击 `称重单整理工具_重做版.exe`。
