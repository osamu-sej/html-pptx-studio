# Mac mini で動かし Windows から使う

Mac mini をサーバーにし、同じ Tailscale ネットワークに参加した Windows のブラウザから使う構成です。アプリは Mac のローカルに待ち受け、Tailscale Serve が HTTPS で中継します。Mac mini は起動・ログインした状態で、スリープせずネットワークにつながっている必要があります。

## Mac mini の準備

Node.js 22 以上、Python 3.11 以上、Google Chrome、Tailscale を用意し、このリポジトリを Mac に取得します。以下はリポジトリのディレクトリで実行します。

```bash
npm ci
python3 -m venv .venv
.venv/bin/pip install -r tools/requirements.txt imageio-ffmpeg
```

`imageio-ffmpeg` は動画を使う PowerPoint 出力用です。次のコマンドで、アプリを Mac のログイン時に起動する LaunchAgent を作成します。使用中の Node、仮想環境内の ffmpeg と Chrome の場所を記録します。既存の同名ファイルがある場合は上書きせず停止します。

```bash
python3 - <<'PY'
import os, pathlib, plistlib, shutil, subprocess
root = pathlib.Path.cwd()
node = shutil.which('node')
chrome = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
python = root / '.venv/bin/python'
if not node or not pathlib.Path(chrome).is_file() or not python.is_file():
    raise SystemExit('Node.js、Google Chrome、または .venv が見つかりません')
ffmpeg = subprocess.check_output([str(python), '-c',
    'import imageio_ffmpeg; print(imageio_ffmpeg.get_ffmpeg_exe())'], text=True).strip()
target = pathlib.Path.home() / 'Library/LaunchAgents/com.osamu.html-pptx-studio.plist'
config = {
    'Label': 'com.osamu.html-pptx-studio',
    'ProgramArguments': [node, str(root / 'server.js')],
    'WorkingDirectory': str(root), 'RunAtLoad': True, 'KeepAlive': True,
    'EnvironmentVariables': {
        'HOST': '127.0.0.1', 'PORT': '8787', 'CHROME_PATH': chrome,
        'FFMPEG_PATH': ffmpeg, 'PPTX_LEAN': '0',
        'PATH': os.path.dirname(node) + ':/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin',
    },
    'StandardOutPath': str(pathlib.Path.home() / 'Library/Logs/html-pptx-studio.log'),
    'StandardErrorPath': str(pathlib.Path.home() / 'Library/Logs/html-pptx-studio-error.log'),
}
target.parent.mkdir(parents=True, exist_ok=True)
with target.open('xb') as f:
    plistlib.dump(config, f)
print('作成:', target)
PY
launchctl bootstrap "gui/$(id -u)" "$HOME/Library/LaunchAgents/com.osamu.html-pptx-studio.plist"
curl -f http://127.0.0.1:8787/healthz
```

`PPTX_LEAN=0` は動画を使う動きの出力を有効にします。動画の撮影には時間がかかる場合があります。自動起動の停止は `launchctl bootout "gui/$(id -u)" "$HOME/Library/LaunchAgents/com.osamu.html-pptx-studio.plist"`、変更後の再起動は `launchctl kickstart -k "gui/$(id -u)/com.osamu.html-pptx-studio"` です。Node の更新やリポジトリの移動をした場合は plist 内のパスも更新してください。

## Tailscale で接続

Mac と Windows の両方で Tailscale にログインします。Mac で `tailscale serve status` を確認し、8788 が別のサービスに使われていなければ次を実行します。既存の Serve や Funnel の設定は削除しないでください。

```bash
tailscale serve --bg --https=8788 http://127.0.0.1:8787
tailscale serve status
```

表示された **tailnet only** の HTTPS URL を、Windows のブラウザで開きます。Tailscale のネットワークに参加できる端末にはアプリが見えるため、参加端末とアクセス権を管理してください。利用をやめるときは `tailscale serve --https=8788 off` でこのポートだけを解除します。

Windows のブラウザで「Codexに接続」を選んで認証します。保存庫や自動保存は使用したブラウザに保存され、Mac と Windows の間で自動共有されません。資料は JSON または HTML で書き出して保管できます。PowerPoint 出力は Windows から開始してダウンロードでき、処理は Mac で行われます。

設定後は Windows から画面を開き、資料の生成・PowerPoint 出力・PowerPoint での再生を確認してください。Mac 側の確認だけでは Windows 側の接続や再生は確認できません。
