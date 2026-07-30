"""通过GitHub API直接推送文件到仓库（绕过git push网络问题）"""
import requests, base64, json, sys, os, io, argparse, builtins
_orig_print = builtins.print
def _safe_print(*args, **kwargs):
    try:
        _orig_print(*args, **kwargs)
    except UnicodeEncodeError:
        safe_args = [str(a).encode('gbk', 'replace').decode('gbk', 'replace') for a in args]
        _orig_print(*safe_args, **kwargs)
builtins.print = _safe_print

def get_token():
    if len(sys.argv) > 1 and not sys.argv[1].startswith('-'):
        return sys.argv[1]
    token_file = os.path.join(os.path.dirname(__file__), '.github_token')
    if os.path.exists(token_file):
        return open(token_file).read().strip()
    return ""

DEFAULT_OWNER = "YHfund1"
DEFAULT_REPO = "jgxw"
DEFAULT_FILES = ["index.html", "data.json"]

parser = argparse.ArgumentParser(description='Deploy files to GitHub Pages')
parser.add_argument('--repo', default=None, help='Owner/Repo, e.g. rilkezhang/jgxw')
parser.add_argument('--files', nargs='+', default=None, help='Files to upload')
parser.add_argument('token', nargs='?', default=None, help='GitHub token (optional)')
args = parser.parse_args()

if args.repo:
    parts = args.repo.split('/')
    OWNER = parts[0]
    REPO = parts[1]
else:
    OWNER = DEFAULT_OWNER
    REPO = DEFAULT_REPO

FILES = args.files if args.files else DEFAULT_FILES
BASE = f"https://api.github.com/repos/{OWNER}/{REPO}"
HEADERS = None

def init_headers(token):
    global HEADERS
    HEADERS = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github+json"
    }

def upload_files():
    blobs = []

    # 1. 创建blob
    for f in FILES:
        if not os.path.exists(f):
            print(f"  ⚠ 文件不存在: {f}")
            return False
        print(f"  上传 {f}...")
        with open(f, "rb") as fh:
            content = base64.b64encode(fh.read()).decode()
        resp = requests.post(f"{BASE}/git/blobs", headers=HEADERS, json={
            "content": content, "encoding": "base64"
        })
        if resp.status_code not in (200, 201):
            print(f"  错误: {resp.status_code} {resp.text[:200]}")
            return False
        blobs.append({"path": f, "mode": "100644", "type": "blob", "sha": resp.json()["sha"]})
        print(f"  ✓ {f} ({len(content)//1024}KB)")

    # 2. 获取当前main分支ref
    print("  获取分支引用...")
    resp = requests.get(f"{BASE}/git/ref/heads/main", headers=HEADERS)
    if resp.status_code != 200:
        print(f"  错误: {resp.status_code} {resp.text[:200]}")
        return False
    old_sha = resp.json()["object"]["sha"]  # commit SHA

    # 3. 从commit中获取tree SHA
    resp = requests.get(f"{BASE}/git/commits/{old_sha}", headers=HEADERS)
    if resp.status_code != 200:
        print(f"  错误: {resp.status_code} {resp.text[:200]}")
        return False
    base_tree_sha = resp.json()["tree"]["sha"]

    # 4. 创建tree
    print("  创建树...")
    resp = requests.post(f"{BASE}/git/trees", headers=HEADERS, json={
        "base_tree": base_tree_sha, "tree": blobs
    })
    if resp.status_code not in (200, 201):
        print(f"  错误: {resp.status_code} {resp.text[:200]}")
        return False
    tree_sha = resp.json()["sha"]

    # 4. 创建commit
    print("  创建提交...")
    file_summary = ', '.join(FILES)
    resp = requests.post(f"{BASE}/git/commits", headers=HEADERS, json={
        "message": f"Deploy: {file_summary}",
        "tree": tree_sha,
        "parents": [old_sha]
    })
    if resp.status_code not in (200, 201):
        print(f"  错误: {resp.status_code} {resp.text[:200]}")
        return False
    commit_sha = resp.json()["sha"]

    # 5. 更新ref
    print("  更新分支...")
    resp = requests.patch(f"{BASE}/git/refs/heads/main", headers=HEADERS, json={
        "sha": commit_sha
    })
    if resp.status_code != 200:
        print(f"  错误: {resp.status_code} {resp.text[:200]}")
        return False

    print("  ✓ 推送成功！")
    return True

if __name__ == "__main__":
    TOKEN = args.token if args.token else get_token()
    if not TOKEN:
        print("用法: python github_deploy.py [TOKEN]")
        print("      python github_deploy.py --repo rilkezhang/jgxw --files ultra_long.html ultra_long_data.json")
        print("或在 .github_token 文件中存放token")
        sys.exit(1)
    init_headers(TOKEN)
    print(f"目标: {OWNER}/{REPO}")
    print(f"文件: {FILES}")
    print("正在通过GitHub API推送文件...")
    if upload_files():
        print("\n推送完成！")
    else:
        print("\n推送失败")
        sys.exit(1)
