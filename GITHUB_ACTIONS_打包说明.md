# GitHub Actions macOS 打包说明

## 要上传到 GitHub 的文件

请上传这些文件和文件夹：

- `.github/`
- `models/`
- `collect_douyin.py`
- `build_release.py`
- `requirements.txt`
- `采集配置.txt`
- `.gitignore`
- `GITHUB_ACTIONS_打包说明.md`

不要上传这些本地数据或打包产物：

- `build/`
- `dist/`
- `release/`
- `screenshots/`
- `采集结果.xlsx`
- `待采集视频链接.xlsx`
- `抖音采集工具_可发送.zip`

## 打包步骤

1. 新建一个 GitHub 仓库，把上面的文件上传进去。
2. 打开仓库页面的 `Actions`。
3. 选择 `Build macOS package`。
4. 点击 `Run workflow`。
5. 等两个任务完成后，在页面底部 `Artifacts` 下载：
   - `抖音采集工具_mac-arm64`：给 Apple Silicon 芯片的 Mac 用，常见于 M1/M2/M3/M4。
   - `抖音采集工具_mac-intel`：给旧款 Intel 芯片的 Mac 用。

如果不知道对方电脑是哪种芯片，就把两个包都发过去。

## Mac 用户第一次打开

Mac 版没有苹果开发者签名，第一次打开可能会被系统拦截。

让对方解压后右键点击 `启动.command`，选择“打开”。如果仍提示无法验证开发者，去“系统设置 > 隐私与安全性”里允许打开，然后再运行。
