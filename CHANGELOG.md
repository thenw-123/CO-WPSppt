# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/) 精神，版本号与 [extension/package.json](extension/package.json) / [manifest.json](manifest.json) 对齐。

## [1.6.0] - 2026-05-03

### Added

- 叙事版式：`thesis-vertical`（竖向三步）、`swot`（四象限）；`Wps.LayoutNarrative.ps1` 与校验、Schema 联动。
- CO-WPSppt 扩展品牌化、`releases/` 分发说明、GitHub Actions Release 工作流。
- 企业向治理与文档：`CONTRIBUTING.md`、`SECURITY.md`、`CHANGELOG.md`、`specs/README.md`、`docs/README.md`、`examples/` 规范、`.editorconfig`、`.gitattributes`。

### Changed

- 完整演示稿与试跑素材迁入 `examples/`（`decks/`、`assets/`、`scripts/`），与 `specs/` 中轻量契约示例分离。
- README 与《GitHub 发布》文档中的本地路径改为 `<仓库根目录>` 占位符；补充 `logs/.gitkeep` 以便忽略下的会话日志仍保留目录结构。
- MIT 版权声明与扩展 `author` 字段统一为 **then_132423@qq.com**（根目录与 `extension/LICENSE` 一致）；`SECURITY.md` 补充同邮箱的私下报告渠道。

## [1.5.x] 及更早

见 Git 历史与 tag；本文件自 1.6.0 起维护。
