import { Config } from '@remotion/cli/config';
import * as fs from 'fs';

// このコンテナ専用のブラウザ指定(ローカルPC・GitHub ActionsではRemotionが自動取得)
const containerShell = '/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell';
if (fs.existsSync(containerShell)) {
  Config.setBrowserExecutable(containerShell);
  Config.setChromiumOpenGlRenderer('angle-egl');
}
Config.setVideoImageFormat('jpeg');
Config.setOverwriteOutput(true);
