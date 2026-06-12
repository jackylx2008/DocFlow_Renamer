# Local AI Runtime Setup

本文档记录当前项目的本地 AI 运行环境：CUDA 基础检查、Python 虚拟环境、`llama.cpp` 服务启动，以及通过根目录脚本验证本地模型调用。

## 当前项目约定

- Python 环境：项目根目录下的 `.venv`
- 配置入口：`config.yaml`
- 本机私有配置：`common.env`
- 本地模型服务：`llama-server.exe`
- API 形式：OpenAI 兼容 HTTP API
- 默认 API 地址：`http://127.0.0.1:8080/v1`
- 自检入口脚本：`ai_self_check.py`
- CUDA 12 runtime DLL：项目内 `vendor/cuda12/`

代码结构遵循 `COMMON_PROJECT_SKILLS.md`：

- `src/localai/modules/`：基础能力模块
- `src/localai/flows/`：场景编排层
- 项目根目录 `.py` 文件：独立入口脚本

## 1. CUDA 基础检查

在项目根目录执行：

```powershell
nvidia-smi
```

正常情况下应看到：

- NVIDIA 显卡名称
- Driver Version
- CUDA Version
- 显存占用信息

这个检查确认驱动与 CUDA 运行能力可被系统识别。当前项目调用 `llama.cpp` 的 CUDA 能力，主要依赖：

- 已编译好的 `llama-server.exe`
- `llama-server.exe` 同目录下的 `ggml-cuda.dll`
- 项目内 `vendor/cuda12/` 下的 CUDA runtime DLL

当前项目随本地目录携带的 CUDA runtime DLL：

```text
vendor/cuda12/cudart64_12.dll
vendor/cuda12/cublas64_12.dll
vendor/cuda12/cublasLt64_12.dll
```

这些 DLL 用于让 `ggml-cuda.dll` 在目标机器上能被 Windows 正确加载。部署到其他电脑时，仍需要目标电脑有可用的 NVIDIA 驱动；项目内 DLL 不替代显卡驱动。

### CUDA runtime DLL 来源

`vendor/cuda12/` 不提交到 git，需要在每台机器本地准备。优先从 NVIDIA 官方 CUDA Toolkit 获取，不要从第三方 DLL 下载站复制。

官方下载入口：

```text
https://developer.nvidia.com/cuda-toolkit-archive
```

推荐做法：

1. 在 NVIDIA CUDA Toolkit Archive 下载与 `llama.cpp` 构建匹配的 CUDA 12.x Windows 安装包。
2. 安装 CUDA Toolkit，或用支持的解压工具从安装包中取出运行时文件。
3. 从 CUDA Toolkit 安装目录复制以下 DLL 到项目目录 `vendor/cuda12/`：

```text
cudart64_12.dll
cublas64_12.dll
cublasLt64_12.dll
```

常见安装位置类似：

```text
C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v12.x\bin\
```

说明：

- `cudart64_12.dll` 属于 CUDA Runtime。
- `cublas64_12.dll` 和 `cublasLt64_12.dll` 属于 cuBLAS runtime。
- DLL 主版本要和 `llama-server.exe` / `ggml-cuda.dll` 构建时使用的 CUDA 主版本一致；当前目录名 `cuda12` 表示使用 CUDA 12 系列 DLL。
- 这些 DLL 只补足本地运行时依赖，不替代 NVIDIA 显卡驱动；驱动仍通过 NVIDIA 官方驱动安装。

## 2. Python 环境

创建虚拟环境：

```powershell
python -m venv .venv
```

安装依赖：

```powershell
.\.venv\Scripts\python.exe -m pip install -U pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

当前 Python 依赖保持很轻，只需要项目配置读取等基础能力；本地模型推理通过 `llama-server` 的 HTTP API 完成，不在 Python 进程内加载 GGUF。

## 3. 本机配置

真实机器路径写在 `common.env`，该文件不入库；仓库只保留 `common.env.example`。

关键变量：

```dotenv
LLAMACPP_BASE_URL=http://127.0.0.1:8080/v1
LLAMACPP_MODEL=local-model
LLAMACPP_AUTOSTART=true

LLAMACPP_SERVER_PATH=
LLAMACPP_MODEL_PATH=
LLAMACPP_MMPROJ_PATH=
LLAMACPP_EXTRA_DLL_DIRS=./vendor/cuda12
LLAMACPP_N_GPU_LAYERS=999
LLAMACPP_CTX_SIZE=8192

LLAMACPP_REASONING=off
LLAMACPP_REASONING_BUDGET=0
```

说明：

- `LLAMACPP_SERVER_PATH` 指向已编译好的 `llama-server.exe`
- `LLAMACPP_MODEL_PATH` 指向主模型 `.gguf`
- `LLAMACPP_MMPROJ_PATH` 指向多模态投影文件；纯文本场景可留空
- `LLAMACPP_EXTRA_DLL_DIRS` 指向额外 DLL 搜索目录；当前默认使用项目内 `./vendor/cuda12`
- `LLAMACPP_MODEL` 会作为 `llama-server --alias` 传入，Python 请求使用这个稳定模型名
- `LLAMACPP_N_GPU_LAYERS=999` 让 `llama.cpp` 尽量将模型层 offload 到显存
- `LLAMACPP_CTX_SIZE=8192` 限制上下文窗口，避免默认超大上下文导致 KV cache 过大
- `LLAMACPP_REASONING=off` 和 `LLAMACPP_REASONING_BUDGET=0` 用于让 Qwen3 类模型在自检时直接返回 `message.content`

当前已验证的本机配置示例：

```dotenv
LLAMACPP_BASE_URL=http://127.0.0.1:8080/v1
LLAMACPP_MODEL=Qwen3.6-27B-Q4_K_M
LLAMACPP_AUTOSTART=true
LLAMACPP_SERVER_PATH=D:\llama-cpp-cu12\llama-server.exe
LLAMACPP_MODEL_PATH=C:\Users\bcjt_\.lmstudio\models\lmstudio-community\Qwen3.6-27B-GGUF\Qwen3.6-27B-Q4_K_M.gguf
LLAMACPP_MMPROJ_PATH=C:\Users\bcjt_\.lmstudio\models\lmstudio-community\Qwen3.6-27B-GGUF\mmproj-Qwen3.6-27B-BF16.gguf
LLAMACPP_EXTRA_DLL_DIRS=./vendor/cuda12
LLAMACPP_N_GPU_LAYERS=999
LLAMACPP_CTX_SIZE=8192
LLAMACPP_REASONING=off
LLAMACPP_REASONING_BUDGET=0
```

## 4. 自动启动流程

运行入口脚本时，项目会执行以下流程：

1. 读取 `common.env`
2. 读取并解析 `config.yaml`
3. 检查 `nvidia-smi`
4. 请求 `GET /health`
5. 请求 `GET /v1/models`
6. 如果服务不可用且 `LLAMACPP_AUTOSTART=true`，自动启动 `llama-server`
7. 轮询等待服务可用
8. 校验配置模型名是否存在于模型列表
9. 可选调用 `POST /v1/chat/completions`

自动启动命令由配置生成，核心参数等价于：

```powershell
& $env:LLAMACPP_SERVER_PATH `
  -m $env:LLAMACPP_MODEL_PATH `
  --mmproj $env:LLAMACPP_MMPROJ_PATH `
  --alias $env:LLAMACPP_MODEL `
  -c $env:LLAMACPP_CTX_SIZE `
  -ngl $env:LLAMACPP_N_GPU_LAYERS `
  --reasoning $env:LLAMACPP_REASONING `
  --reasoning-budget $env:LLAMACPP_REASONING_BUDGET `
  --host 127.0.0.1 `
  --port 8080 `
  --verbose
```

如果 `LLAMACPP_MMPROJ_PATH` 为空，启动命令不会传 `--mmproj`。

自动启动时，`LlamaCppClient` 会把以下目录加入 `llama-server.exe` 子进程的 `PATH`：

1. `llama-server.exe` 所在目录
2. `LLAMACPP_EXTRA_DLL_DIRS` 中配置的目录

因此 `LLAMACPP_EXTRA_DLL_DIRS=./vendor/cuda12` 会被解析为项目根目录下的绝对路径，再传给子进程。这样 `ggml-cuda.dll` 能找到 CUDA runtime DLL。

## 5. CUDA backend 验证

先确认 `llama.cpp` 能看到 CUDA 设备：

```powershell
$env:PATH=(Resolve-Path ".\vendor\cuda12").Path + ";D:\llama-cpp-cu12;" + $env:PATH
& "D:\llama-cpp-cu12\llama-server.exe" --list-devices
```

正常情况下应看到类似：

```text
ggml_cuda_init: found 1 CUDA devices
load_backend: loaded CUDA backend from D:\llama-cpp-cu12\ggml-cuda.dll
Available devices:
  CUDA0: NVIDIA GeForce RTX 5090 D v2
```

如果只看到：

```text
load_backend: loaded CPU backend ...
Available devices:
```

说明 CUDA backend 没有正确加载。优先检查：

- `vendor/cuda12/` 下是否存在 `cudart64_12.dll`、`cublas64_12.dll`、`cublasLt64_12.dll`
- `LLAMACPP_EXTRA_DLL_DIRS` 是否指向 `./vendor/cuda12`
- `llama-server.exe` 同目录是否存在 `ggml-cuda.dll`
- NVIDIA 驱动是否可被 `nvidia-smi` 识别

## 6. 项目自检

只检查 CUDA、服务健康状态和模型列表，不发送对话请求：

```powershell
.\.venv\Scripts\python.exe ai_self_check.py --no-chat
```

完整端到端检查：

```powershell
.\.venv\Scripts\python.exe ai_self_check.py --max-tokens 64 --prompt "请直接回答两个字：可用"
```

成功时会输出类似：

```json
{
  "cuda_check": {
    "command": "nvidia-smi",
    "ok": true
  },
  "llamacpp": {
    "base_url": "http://127.0.0.1:8080/v1",
    "model": "local-model",
    "health": {
      "status": "ok"
    },
    "available_models": [
      "local-model"
    ]
  },
  "answer": "可用"
}
```

最小本地 AI ping 测试：

```powershell
.\.venv\Scripts\python.exe ai_self_check.py --prompt "你好" --max-tokens 32
```

复杂一点的生成测试：

```powershell
.\.venv\Scripts\python.exe ai_self_check.py --prompt "用C语言写一个冒泡排序法" --max-tokens 128
```

当前约定：测试脚本和生产流程在本次运行结束时都会关闭本次启动的 `llama-server.exe`，释放显存。

## 7. 手动接口检查

服务启动后，可以直接检查：

```powershell
Invoke-RestMethod http://127.0.0.1:8080/health
Invoke-RestMethod http://127.0.0.1:8080/v1/models
```

手动发送一次 OpenAI 兼容请求：

```powershell
$body = @{
  model = $env:LLAMACPP_MODEL
  temperature = 0
  max_tokens = 64
  messages = @(
    @{
      role = "user"
      content = "请直接回答两个字：可用"
    }
  )
} | ConvertTo-Json -Depth 10

Invoke-RestMethod `
  -Method Post `
  -Uri "http://127.0.0.1:8080/v1/chat/completions" `
  -ContentType "application/json" `
  -Body $body
```

## 8. 日志与排错

项目日志：

```text
log/<entry_name>.log
```

`llama-server` 输出：

```text
log/llama_server.out.log
log/llama_server.err.log
```

常见问题：

- `nvidia-smi` 不可用：先检查 NVIDIA 驱动是否正常安装
- `/health` 不可用：检查端口是否被占用，或 `LLAMACPP_SERVER_PATH` 是否正确
- 模型不存在：检查 `LLAMACPP_MODEL_PATH`，并确认 `.gguf` 文件存在
- 请求模型名不匹配：访问 `/v1/models`，将 `LLAMACPP_MODEL` 改为返回的模型 id，或使用 `--alias` 保持稳定名称
- 只返回 reasoning、不返回 content：确认 `LLAMACPP_REASONING=off` 和 `LLAMACPP_REASONING_BUDGET=0`，然后重启 `llama-server`
- 中文输出乱码：使用项目入口脚本输出；脚本已将 stdout 设置为 UTF-8
- `llama-server.exe` 进程内存高但显存低：先用 `--list-devices` 确认 CUDA backend 是否加载；若未加载，检查 `vendor/cuda12` 和 `LLAMACPP_EXTRA_DLL_DIRS`
- 任务结束后显存未释放：检查是否仍有 `llama-server.exe` 残留进程

停止本地服务：

```powershell
Get-Process llama-server -ErrorAction SilentlyContinue | Stop-Process
```

## 9. 进程生命周期与显存释放

项目当前统一要求：测试代码和生产代码结束时释放显存，关闭本次启动的 `llama-server.exe`。

实现约定：

- 通过 `LlamaCppClient.ensure_server()` 自启动的进程会记录在 client 内部
- 入口或 flow 必须使用 `try/finally`
- `finally` 中调用 `client.shutdown_server()`
- 如果连接的是外部已存在服务，`shutdown_server()` 不会关闭它，因为该进程不是当前 client 启动的

典型写法：

```python
client = LlamaCppClient(llama_config)
try:
    health, models = client.ensure_server()
    client.assert_model_available(models)
    answer = client.chat("你好", max_tokens=32)
finally:
    client.shutdown_server()
```

`ai_self_check.py` 已按此方式处理。后续订单截图 OCR、订单结构化识别入口也应沿用同样的生命周期管理方式。

## 10. 新入口脚本接入方式

新增一个根目录入口脚本时，应保持入口层很薄：

1. 调用 `bootstrap_context(__file__)`
2. 初始化配置、日志和上下文
3. 调用 `src/localai/flows/` 中对应的 `run(...)`
4. 用 `print_json(...)` 输出结构化结果

需要调用本地模型时，优先在 `flows` 中组合：

- `LlamaCppConfig.from_config(...)`
- `LlamaCppClient.ensure_server()`
- `LlamaCppClient.assert_model_available(...)`
- `LlamaCppClient.chat(...)`
- `LlamaCppClient.shutdown_server()`

不要在入口脚本里直接拼接模型路径、启动命令或 HTTP 请求。
