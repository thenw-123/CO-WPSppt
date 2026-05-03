# 主题（.thmx）目录（L2）

将 WPS / PowerPoint 兼容的 **`.thmx`** 文件放在本目录（或子目录），在 spec 里写相对路径即可：

```json
"theme": {
  "themePath": "themes/your-brand.thmx",
  "titleFont": "Microsoft YaHei",
  "bodyFont": "Microsoft YaHei"
}
```

- 仓库不强制附带 `.thmx`（体积与版权原因）；请从企业模板或 WPS 导出自有主题。
- 若路径无效，`ApplyTheme` 会静默跳过，字体仍可按 `titleFont`/`bodyFont` 尽量套用。
