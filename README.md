# FakeWordMfc

## 项目简介

FakeWordMfc 是一个基于 MFC 的 Windows 桌面应用程序，用于学习/工作中摸鱼，使用该软件，你可以：
- 在Word中假装打字从而看小说
- 当老板来时使用快捷键切换到工作模式，从而假装认真工作

## 主要模块

- **FakeWordMfc**：主程序，包含窗口界面、配置管理、Word 控制等功能。
- **WordDocumentInputHook**：输入钩子 DLL，用于拦截和处理 Word 文档输入事件。

## 编译环境

- Visual Studio 2022 或更高版本
- Windows 10 SDK
- MFC 库
- 平台工具集 v143

## 运行方式

1. 使用 Visual Studio 打开 `FakeWordMfc.sln`，编译解决方案。
2. 运行 `FakeWordMfc.exe`，如需钩子功能请确保 `WordDocumentInputHook.dll` 在同目录下。
3. 配置文件位于 `x64/Debug/config.ini`，可根据需要修改。

## 目录结构

- `FakeWordMfc/` 主程序源码
- `WordDocumentInputHook/` 钩子 DLL 源码


## 许可证

详见 LICENSE。
