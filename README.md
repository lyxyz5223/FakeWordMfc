# FakeWordMfc

## 简介

FakeWordMfc 是一个基于 MFC 的 Windows 桌面应用程序，用于学习/工作中摸鱼，使用该软件，你可以：
- 在Word中假装打字从而看小说
- 当老板来时使用快捷键切换到工作模式，从而假装认真工作

## 使用方法：

- 打开软件
- 选择文件/输入文件路径，并非摸鱼和工作文件都需要选择，使用什么模式就选择对应文件即可，如需老板键躲检查，可以同时选择摸鱼和工作文件
- 点击“打开并读取文件”按钮
- 可选：点击“设置”按钮打开设置查看/修改当前快捷键等信息
- 点击“打开Word/WPS Office，开始fish”按钮
- 即刻开始你的摸鱼之旅吧！(*^▽^*)

## 特性：

- 支持不同类型的小说和工作文件
- 支持输入时按照一定比例输出小说/工作内容
- 支持状态窗口显示当前状态信息，但由于窗口的生命周期与内存管理问题，暂不支持将状态窗口嵌入Word窗口中
- 支持快捷键一键切换工作/摸鱼模式，一键防检查！
- 支持不同的Word文本改变检测方式

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
