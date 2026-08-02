"""
Qwen2.5 0.5B Hyperparameter Lab

설치:
    py -m pip install -U huggingface-hub llama-cpp-python \
        --extra-index-url https://abetlen.github.io/llama-cpp-python/whl/cpu

실행:
    py qwen_hyperparameter_lab.py

브라우저:
    http://127.0.0.1:8000
"""

from __future__ import annotations

import gc
import json
import math
import os
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any
from urllib.parse import urlparse

MODEL_REPO = "Qwen/Qwen2.5-0.5B-Instruct-GGUF"
MODEL_FILE = "qwen2.5-0.5b-instruct-q4_k_m.gguf"
HOST = "127.0.0.1"
PORT = 8000

DEFAULT_RUNTIME: dict[str, Any] = {
    "n_ctx": 4096,
    "n_batch": 512,
    "n_ubatch": 512,
    "n_threads": max(1, (os.cpu_count() or 4) // 2),
    "n_threads_batch": max(1, os.cpu_count() or 4),
    "n_gpu_layers": 0,
    "last_n_tokens_size": 64,
    "use_mmap": True,
    "use_mlock": False,
    "offload_kqv": True,
    "flash_attn": False,
    "op_offload": None,
    "swa_full": None,
    "numa": False,
    "rope_freq_base": 0.0,
    "rope_freq_scale": 0.0,
    "yarn_ext_factor": -1.0,
    "yarn_attn_factor": 1.0,
    "yarn_beta_fast": 32.0,
    "yarn_beta_slow": 1.0,
    "yarn_orig_ctx": 0,
}

MODEL: Any = None
GRAMMAR_CLASS: Any = None
MODEL_LOCK = threading.RLock()
RUNTIME = dict(DEFAULT_RUNTIME)
STATUS = {"ready": False, "loading": False, "error": None}

HTML = r'''<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Qwen Hyperparameter Lab</title>
<style>
:root{color-scheme:light dark;--bg:#f4f6fa;--p:#fff;--p2:#f0f3f8;--t:#172033;--m:#667085;--b:#d7dde7;--a:#315efb;--as:#e8eeff;--e:#b42318;--w:#b54708;--ok:#067647}
@media(prefers-color-scheme:dark){:root{--bg:#0e1117;--p:#171b23;--p2:#202631;--t:#e7ebf2;--m:#98a2b3;--b:#303846;--a:#7894ff;--as:#263150;--e:#ff8a80;--w:#ffb86b;--ok:#5bd39a}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--t);font-family:Inter,"Noto Sans KR",system-ui,sans-serif}button,input,textarea,select{font:inherit}.app{display:grid;grid-template-columns:350px minmax(430px,1fr)350px;height:100vh;gap:1px;background:var(--b)}.panel{min-width:0;overflow:auto;background:var(--p)}.side,.inspect{padding:18px}h1{font-size:18px;margin:0 0 4px}h2{font-size:15px;margin:0}h3{font-size:13px}.small{font-size:12px;line-height:1.5;color:var(--m)}.tabs,.actions{display:grid;grid-template-columns:1fr 1fr;gap:7px}.tabs{margin:16px 0}.btn,.tab{border:1px solid var(--b);border-radius:9px;padding:9px;background:var(--p2);color:var(--t);cursor:pointer}.tab.on,.btn.primary{border-color:var(--a);background:var(--a);color:#fff}.btn:disabled{opacity:.55;cursor:not-allowed}.full{width:100%}.hide{display:none}details{border-top:1px solid var(--b);padding:13px 0}summary{font-size:13px;font-weight:650;cursor:pointer;margin-bottom:11px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.field.all{grid-column:1/-1}.field label{display:flex;align-items:center;gap:6px;margin-bottom:5px;font-size:12px}.field input,.field textarea,.field select{width:100%;border:1px solid var(--b);border-radius:8px;padding:8px;background:var(--p2);color:var(--t)}textarea{min-height:72px;resize:vertical}.check{display:flex;align-items:center;gap:7px;min-height:35px;font-size:12px}.check input{width:auto}.help{margin-top:4px;font-size:11px;color:var(--m);line-height:1.4}.q{width:19px;height:19px;padding:0;border:1px solid var(--b);border-radius:50%;background:var(--p2);color:var(--m);font-size:11px;font-weight:800;cursor:pointer}.q:hover{border-color:var(--a);color:var(--a)}.note{margin:10px 0;padding:10px;border:1px solid var(--b);border-radius:9px;background:var(--as);font-size:12px;line-height:1.5}.warn{color:var(--w);background:var(--p2)}
.chat{display:grid;grid-template-rows:auto 1fr auto;overflow:hidden}.top{display:flex;justify-content:space-between;align-items:center;padding:13px 17px;border-bottom:1px solid var(--b)}.status{display:flex;align-items:center;gap:7px;font-size:12px;color:var(--m)}.dot{width:8px;height:8px;border-radius:50%;background:var(--ok)}.dot.load{background:var(--w)}.dot.err{background:var(--e)}.messages{overflow:auto;padding:22px}.empty{max-width:600px;margin:12vh auto 0;text-align:center;color:var(--m)}.msg{max-width:850px;margin:0 auto 17px}.role{font-size:11px;color:var(--m);font-weight:700;margin-bottom:5px}.bubble{padding:13px 15px;border:1px solid var(--b);border-radius:13px;white-space:pre-wrap;overflow-wrap:anywhere;line-height:1.7}.user .bubble{background:var(--a);color:#fff;border-color:transparent}.tok{border-radius:3px;cursor:pointer}.tok:hover,.tok.sel{background:var(--as);outline:1px solid var(--a)}.tok.sel{outline-width:2px}.compose{padding:13px 17px 17px;border-top:1px solid var(--b)}.compose>div:first-child{max-width:880px;margin:auto;display:grid;grid-template-columns:1fr auto;gap:9px;align-items:end}#prompt{min-height:55px;max-height:180px;width:100%;border:1px solid var(--b);border-radius:9px;padding:9px;background:var(--p2);color:var(--t)}.error{max-width:880px;margin:7px auto 0;color:var(--e);font-size:12px}
.tokenTitle{margin:13px 0;padding:11px;border-radius:8px;background:var(--p2);font:17px ui-monospace,Consolas,monospace;white-space:pre-wrap}.stats{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:16px}.stat{padding:9px;border-radius:8px;background:var(--p2)}.cand{margin:9px 0}.candHead{display:flex;justify-content:space-between;gap:8px;font-size:12px;margin-bottom:4px}.candName{max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-family:ui-monospace,Consolas,monospace}.chosen{color:var(--a);font-weight:700}.bar{height:7px;border-radius:99px;background:var(--p2);overflow:hidden}.bar div{height:100%;min-width:1px;background:var(--a)}
.back{position:fixed;inset:0;z-index:100;display:grid;place-items:center;padding:20px;background:#0009}.back[hidden]{display:none}.modal{width:min(590px,100%);max-height:90vh;overflow:auto;border:1px solid var(--b);border-radius:15px;background:var(--p)}.modalHead{position:sticky;top:0;display:flex;justify-content:space-between;gap:10px;padding:17px;border-bottom:1px solid var(--b);background:var(--p)}.close{width:32px;height:32px;border:1px solid var(--b);border-radius:9px;background:var(--p2);color:var(--t);cursor:pointer}.modalBody{padding:18px}.facts{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin:15px 0}.fact{padding:10px;border-radius:9px;background:var(--p2)}.fact b{display:block;font-size:12px;margin-bottom:4px}.section{margin-top:15px}.section p{margin:0;font-size:13px;line-height:1.65}.caution{padding:10px;border-left:3px solid var(--w);background:var(--p2)}
@media(max-width:1150px){.app{grid-template-columns:320px 1fr}.inspect{position:fixed;inset:0 0 0 auto;width:min(370px,94vw);z-index:20;border-left:1px solid var(--b);transform:translateX(100%);transition:.2s}.inspect.open{transform:none}}@media(max-width:760px){.app{display:block}.side{display:none}.chat{height:100vh}}
</style>
</head>
<body>
<div class="app">
<aside class="panel side">
<h1>Qwen Hyperparameter Lab</h1><div class="small">Qwen2.5 0.5B Instruct · Q4_K_M</div>
<div class="tabs"><button class="tab on" data-pane="gen">생성 설정</button><button class="tab" data-pane="run">실행 설정</button></div>
<section id="genPane">
<div class="actions"><button id="precise" class="btn">정확성 preset</button><button id="creative" class="btn">창의성 preset</button></div>
<details open><summary>프롬프트</summary><div class="field all"><label>System prompt <button class="q" data-help="system_prompt">?</button></label><textarea id="system_prompt">당신은 정확하고 간결한 AI 어시스턴트입니다. 사용자가 사용한 언어로 답변하세요.</textarea></div></details>
<details open><summary>기본 샘플링</summary><div id="sampling" class="grid"></div><div id="miroNote" class="note warn hide">Mirostat 사용 시 top-k, top-p, min-p, typical-p, TFS는 사용되지 않습니다.</div></details>
<details><summary>Penalty와 재현성</summary><div id="penalty" class="grid"></div></details>
<details><summary>출력 제한</summary><div id="output" class="grid"></div></details>
<details><summary>Mirostat</summary><div id="miro" class="grid"></div></details>
<details><summary>Logit bias</summary><div class="field all"><label>Token ID → bias JSON <button class="q" data-help="logit_bias">?</button></label><textarea id="logit_bias" placeholder='{"198":-5,"151643":-100}'></textarea></div></details>
<details><summary>문법 제약</summary><div class="grid"><div class="field all"><label>Grammar mode <button class="q" data-help="grammar_mode">?</button></label><select id="grammar_mode"><option value="none">사용 안 함</option><option value="gbnf">GBNF</option><option value="json_schema">JSON Schema</option></select></div><div class="field all"><label>Grammar / schema <button class="q" data-help="grammar_text">?</button></label><textarea id="grammar_text" disabled></textarea></div></div></details>
</section>
<section id="runPane" class="hide"><div class="note">실행 설정 적용 시 모델을 다시 로드하고 대화 기록을 초기화합니다.</div><details open><summary>문맥과 배치</summary><div id="context" class="grid"></div></details><details open><summary>CPU와 GPU</summary><div id="device" class="grid"></div></details><details><summary>메모리와 오프로딩</summary><div id="memory" class="grid"></div></details><details><summary>RoPE / YaRN</summary><div id="rope" class="grid"></div><div class="note warn">검증된 모델 설정이 없다면 RoPE는 0, YaRN ext factor는 -1로 유지하십시오.</div></details><button id="reload" class="btn primary full">적용 및 모델 재로드</button></section>
<div style="border-top:1px solid var(--b);padding-top:13px;margin-top:13px"><button id="reset" class="btn full">대화 초기화</button></div>
</aside>
<main class="panel chat"><header class="top"><div><b>Localhost Chat</b><div id="usage" class="small">토큰을 클릭하면 후보 확률을 볼 수 있습니다.</div></div><div class="status"><span id="dot" class="dot"></span><span id="status">확인 중</span></div></header><section id="messages" class="messages"><div id="empty" class="empty"><h2>설정을 조절하고 메시지를 입력하세요.</h2><p>생성 토큰을 클릭하면 해당 위치의 상위 후보와 확률이 표시됩니다.</p></div></section><footer class="compose"><div><textarea id="prompt" placeholder="Enter 전송 · Shift+Enter 줄바꿈"></textarea><button id="send" class="btn primary">전송</button></div><div id="error" class="error"></div></footer></main>
<aside id="inspector" class="panel inspect"><h2>Token inspector</h2><div class="small">선택 토큰의 원시 log probability</div><div id="inspectBody" class="small" style="margin-top:28px">답변에서 토큰을 선택하세요.</div></aside>
</div>
<div id="back" class="back" hidden><section class="modal" role="dialog" aria-modal="true" aria-labelledby="helpTitle"><header class="modalHead"><div><h2 id="helpTitle"></h2><div id="helpGroup" class="small"></div></div><button id="close" class="close">×</button></header><div class="modalBody"><p id="helpSummary" style="margin:0;line-height:1.7"></p><div class="facts"><div class="fact"><b>기본값</b><span id="helpDefault"></span></div><div class="fact"><b>일반 권장 범위</b><span id="helpRange"></span></div></div><div class="section"><h3>값을 변경하면</h3><p id="helpEffect"></p></div><div class="section caution"><h3>주의사항</h3><p id="helpCaution"></p></div></div></section></div>
<script>
const $=id=>document.getElementById(id),history=[];let busy=false,selected=null,lastHelp=null;
const H=(summary,def,range,effect,caution,group)=>({summary,def,range,effect,caution,group});
const HELP={
system_prompt:H("모델이 대화 전체에서 따라야 할 역할·말투·규칙을 system 역할로 전달합니다.","현재 입력 문장","자유 텍스트","구체적일수록 형식과 행동이 안정됩니다.","보안 경계가 아니며 0.5B 모델은 복잡한 규칙을 놓칠 수 있습니다.","입력 지침"),
temperature:H("토큰 확률 분포의 날카로움을 조절합니다.","0.7","0.1–1.2","낮으면 보수적, 높으면 다양하지만 오류 가능성도 증가합니다. 0은 greedy입니다.","top-p·top-k와 함께 과도하게 제한하지 마십시오.","샘플링"),
top_p:H("누적 확률이 P가 될 때까지 높은 확률 후보만 남깁니다.","0.9","0.8–0.95","낮으면 후보가 줄고 안정적이며, 1에 가까우면 후보가 많아집니다.","Mirostat 활성화 시 사용되지 않습니다.","샘플링"),
top_k:H("확률 상위 K개 토큰만 후보로 남깁니다.","40","20–100","작으면 보수적, 크면 드문 후보도 허용합니다. 0은 보통 비활성입니다.","Mirostat 활성화 시 사용되지 않습니다.","샘플링"),
min_p:H("최고 확률 토큰에 비해 너무 낮은 후보를 제거합니다.","0.05","0.02–0.10","높이면 불확실한 후보를 더 강하게 제거합니다.","다른 필터가 강하면 효과가 작고 Mirostat에서는 무시됩니다.","샘플링"),
typical_p:H("정보량이 지나치게 흔하거나 드문 후보를 제외하는 typical sampling입니다.","1.0","0.8–1.0","낮추면 전형적인 정보량의 후보에 집중합니다.","일반 채팅은 1.0 권장, Mirostat에서는 무시됩니다.","샘플링"),
tfs_z:H("확률 분포 꼬리를 잘라내는 Tail-Free Sampling 강도입니다.","1.0","0.8–1.0","1보다 낮으면 후순위 후보를 더 제거합니다.","다른 필터와 중복될 수 있고 Mirostat에서는 무시됩니다.","샘플링"),
repeat_penalty:H("최근 문맥에 등장한 토큰의 재선택을 억제합니다.","1.1","1.0–1.2","높이면 반복이 줄지만 필요한 조사·기호도 억제될 수 있습니다.","1.0은 비활성이고 Penalty history 범위에 적용됩니다.","반복 제어"),
presence_penalty:H("한 번이라도 등장한 토큰에 고정 페널티를 줍니다.","0.0","0–1","양수는 새 어휘·주제로 이동시키고 음수는 재사용을 장려합니다.","repeat penalty와 중복 적용됩니다.","반복 제어"),
frequency_penalty:H("토큰 등장 횟수에 비례해 페널티를 줍니다.","0.0","0–1","높이면 반복 횟수가 많을수록 강하게 억제합니다.","코드·목록처럼 기호 반복이 필요한 출력에서는 불리할 수 있습니다.","반복 제어"),
seed:H("샘플링 난수 생성기의 시작값입니다.","비움","0–2147483647","고정하면 같은 조건에서 결과 재현성이 높아집니다.","라이브러리·백엔드·스레드가 달라지면 결과도 달라질 수 있습니다.","재현성"),
max_tokens:H("응답에서 새로 생성할 최대 토큰 수입니다.","256","64–512","높이면 긴 답변이 가능하지만 시간과 메모리가 증가합니다.","입력+출력 토큰은 Context size를 넘을 수 없습니다.","출력 제한"),
top_logprobs:H("각 위치에서 UI에 표시할 상위 후보 수입니다.","15","5–20","높이면 후보 비교가 늘지만 응답 데이터도 커집니다.","생성 결과 자체에는 영향을 주지 않습니다.","분석 UI"),
response_format:H("일반 텍스트 또는 JSON 객체 출력을 지정합니다.","text","text/json_object","JSON을 선택하면 구조적 출력 제약을 사용합니다.","Grammar와 동시에 사용하지 않도록 코드가 제한합니다.","출력 제약"),
stop:H("지정 문자열이 나타나면 생성을 중단합니다.","없음","문자열 여러 개","역할 구분자 이후 출력 등을 차단할 수 있습니다.","짧고 흔한 문자열은 지나친 조기 종료를 만듭니다.","출력 제한"),
mirostat_mode:H("목표 정보량을 유지하도록 분포를 적응적으로 조절합니다.","0","0/1/2","1 또는 2는 tau를 목표로 다양성을 자동 조절합니다.","활성화 시 top-k·top-p·min-p·typical-p·TFS가 무시됩니다.","적응형 샘플링"),
mirostat_tau:H("Mirostat이 유지하려는 목표 surprisal입니다.","5.0","3–8","높이면 더 다양하고 낮추면 더 예측 가능한 출력을 만듭니다.","Mirostat mode가 0이면 사용되지 않습니다.","적응형 샘플링"),
mirostat_eta:H("Mirostat 목표 오차를 보정하는 학습률입니다.","0.1","0.05–0.2","높으면 빠르게 반응하지만 변동이 커질 수 있습니다.","Mirostat mode가 0이면 사용되지 않습니다.","적응형 샘플링"),
logit_bias:H("특정 tokenizer token ID의 logit을 직접 증감합니다.","없음","-100–100","양수는 장려, 음수는 억제합니다.","텍스트가 아닌 토큰 ID이며 강한 값은 문장을 붕괴시킬 수 있습니다.","토큰 직접 제어"),
grammar_mode:H("허용되는 토큰을 GBNF 또는 JSON Schema 문법에 맞게 제한합니다.","none","none/GBNF/JSON Schema","정해진 구조의 JSON·코드·선택지를 강제할 수 있습니다.","복잡하거나 잘못된 문법은 생성 실패·지연을 유발합니다.","구조적 출력"),
grammar_text:H("GBNF 문법 또는 JSON Schema 본문입니다.","비어 있음","유효한 문법","모델은 문법적으로 허용된 토큰만 생성합니다.","JSON Schema 모드는 유효한 JSON이어야 합니다.","구조적 출력"),
n_ctx:H("시스템 프롬프트·대화·출력을 포함한 최대 문맥 토큰 수입니다.","4096","2048–8192","높이면 긴 대화가 가능하지만 KV 캐시 메모리가 증가합니다.","모델 학습 길이 초과 시 품질이 저하될 수 있습니다.","모델 실행"),
n_batch:H("프롬프트 평가의 논리적 최대 배치 크기입니다.","512","128–1024","높이면 프롬프트 처리가 빨라질 수 있지만 메모리 피크가 증가합니다.","Context size 이하여야 합니다.","모델 실행"),
n_ubatch:H("실제 연산 단위인 physical batch 크기입니다.","512","64–512","낮추면 메모리 피크가 줄고 높이면 병렬성이 좋아질 수 있습니다.","Prompt batch보다 클 수 없습니다.","모델 실행"),
last_n_tokens_size:H("반복 페널티가 참조할 최근 토큰 수입니다.","64","32–256","높이면 오래 전 반복까지 억제합니다.","너무 높으면 필요한 용어 재사용도 막습니다.","반복 범위"),
n_threads:H("토큰 생성 단계에서 사용하는 CPU 스레드 수입니다.","논리 코어 절반","물리 코어 전후","적정 범위까지 높이면 생성 속도가 개선됩니다.","너무 높으면 스레드 경합으로 느려질 수 있습니다.","CPU 실행"),
n_threads_batch:H("프롬프트 배치 평가에 사용하는 CPU 스레드 수입니다.","논리 코어 수","물리–논리 코어 수","긴 입력의 초기 처리 속도에 주로 영향을 줍니다.","높으면 시스템 응답성이 떨어질 수 있습니다.","CPU 실행"),
n_gpu_layers:H("GPU로 오프로딩할 모델 레이어 수입니다. 0=CPU, -1=전체입니다.","0","0 또는 -1부터 실험","GPU 지원 빌드에서는 속도가 크게 개선될 수 있습니다.","VRAM 부족 또는 CPU wheel이면 실패할 수 있습니다.","GPU 실행"),
numa:H("NUMA 시스템에서 메모리 배치를 최적화합니다.","false","다중 소켓 서버만","NUMA 서버에서 메모리 지역성을 개선할 수 있습니다.","일반 PC에서는 이점이 없거나 느려질 수 있습니다.","CPU 메모리"),
use_mmap:H("모델 파일을 운영체제 memory mapping으로 읽습니다.","true","대부분 true","로딩과 OS 페이지 캐시 활용이 효율적입니다.","느린 네트워크 드라이브에서는 불리할 수 있습니다.","메모리 로딩"),
use_mlock:H("모델 메모리가 swap으로 이동하지 않도록 잠급니다.","false","RAM 충분할 때만","swap 지연을 줄일 수 있습니다.","권한이 필요하며 RAM 부족 시 시스템이 불안정해질 수 있습니다.","메모리 로딩"),
offload_kqv:H("Attention의 K/Q/V 연산과 캐시를 GPU로 오프로딩합니다.","true","GPU 사용 시 true","GPU attention 속도를 높이지만 VRAM을 더 사용합니다.","CPU 실행에서는 이점이 제한적입니다.","GPU 실행"),
flash_attn:H("지원 백엔드의 메모리 효율적인 attention 커널을 사용합니다.","false","지원 GPU에서 true","긴 문맥의 속도·메모리 효율이 개선될 수 있습니다.","지원하지 않는 빌드에서는 실패하거나 무시됩니다.","Attention 최적화"),
op_offload:H("호스트 tensor 연산을 장치로 오프로딩합니다.","auto","auto 권장","지원 백엔드에서 CPU/GPU 작업 분배를 바꿉니다.","버전에 따라 지원과 효과가 다릅니다.","GPU 실행"),
swa_full:H("Sliding Window Attention에 전체 크기 캐시를 사용할지 정합니다.","auto","auto 권장","특정 모델의 호환성을 높이지만 메모리가 늘 수 있습니다.","이 Qwen 모델에서는 효과가 제한적일 수 있습니다.","KV 캐시"),
rope_freq_base:H("RoPE 기본 주파수를 재정의합니다. 0은 모델 기본값입니다.","0","0 권장","위치 회전 주파수와 긴 문맥 동작을 바꿉니다.","임의 변경은 품질을 크게 손상시킬 수 있습니다.","위치 인코딩"),
rope_freq_scale:H("RoPE 주파수 스케일을 재정의합니다. 0은 모델 기본값입니다.","0","0 권장","문맥 위치의 압축·확장 방식에 영향을 줍니다.","검증된 scaling 값이 있을 때만 사용하십시오.","위치 인코딩"),
yarn_ext_factor:H("YaRN 문맥 확장 보간 강도입니다. -1은 자동입니다.","-1","-1 권장","원래 학습 길이보다 긴 문맥을 위한 위치 보간에 관여합니다.","검증되지 않은 값은 품질을 급격히 떨어뜨립니다.","문맥 확장"),
yarn_attn_factor:H("YaRN 적용 시 attention 크기 보정 계수입니다.","1","공식 설정 사용","긴 문맥에서 attention 분포 크기를 보정합니다.","다른 YaRN 값과 함께 검증된 설정만 사용하십시오.","문맥 확장"),
yarn_beta_fast:H("YaRN의 빠른 주파수 영역 경계입니다.","32","공식 설정 사용","어떤 RoPE 차원에 강한 보간을 적용할지 바꿉니다.","독립 튜닝보다 기본값 유지가 안전합니다.","문맥 확장"),
yarn_beta_slow:H("YaRN의 느린 주파수 영역 경계입니다.","1","공식 설정 사용","낮은 주파수 위치 성분의 보간 범위를 바꿉니다.","기본값 유지가 안전합니다.","문맥 확장"),
yarn_orig_ctx:H("YaRN 계산 기준이 되는 원래 학습 문맥 길이입니다. 0은 자동입니다.","0","모델 공식 길이","확장 비율 계산의 기준점이 됩니다.","실제 학습 길이와 다르면 위치 스케일링이 잘못됩니다.","문맥 확장")};
const GEN=[
["sampling","temperature","Temperature","number",0.7,-1,2,.05],["sampling","top_p","Top P","number",.9,0,1,.01],["sampling","top_k","Top K","number",40,0,100000,1],["sampling","min_p","Min P","number",.05,0,1,.01],["sampling","typical_p","Typical P","number",1,0,1,.01],["sampling","tfs_z","TFS Z","number",1,0,2,.01],
["penalty","repeat_penalty","Repeat penalty","number",1.1,0,3,.01],["penalty","presence_penalty","Presence penalty","number",0,-2,2,.05],["penalty","frequency_penalty","Frequency penalty","number",0,-2,2,.05],["penalty","seed","Seed","number","",0,2147483647,1],
["output","max_tokens","Max tokens","number",256,1,4096,1],["output","top_logprobs","표시 후보 수","number",15,1,50,1],["output","response_format","Response format","select","text",[["text","일반 텍스트"],["json_object","JSON object"]]],["output","stop","Stop 문자열","textarea",""],
["miro","mirostat_mode","Mirostat mode","select","0",[["0","사용 안 함"],["1","v1"],["2","v2"]]],["miro","mirostat_tau","Mirostat tau","number",5,0,20,.1],["miro","mirostat_eta","Mirostat eta","number",.1,0,1,.01]];
const RUN=[
["context","n_ctx","Context size","number",4096,256,32768,256],["context","n_batch","Prompt batch","number",512,1,4096,1],["context","n_ubatch","Physical batch","number",512,1,4096,1],["context","last_n_tokens_size","Penalty history","number",64,0,4096,1],
["device","n_threads","Generation threads","number",Math.max(1,Math.floor((navigator.hardwareConcurrency||4)/2)),1,256,1],["device","n_threads_batch","Batch threads","number",navigator.hardwareConcurrency||4,1,256,1],["device","n_gpu_layers","GPU layers","number",0,-1,200,1],["device","numa","NUMA","check",false],
["memory","use_mmap","Use mmap","check",true],["memory","use_mlock","Use mlock","check",false],["memory","offload_kqv","Offload K/Q/V","check",true],["memory","flash_attn","Flash attention","check",false],["memory","op_offload","Operation offload","select","auto",[["auto","자동"],["true","사용"],["false","사용 안 함"]]],["memory","swa_full","Full SWA cache","select","auto",[["auto","자동"],["true","사용"],["false","사용 안 함"]]],
["rope","rope_freq_base","RoPE frequency base","number",0,-1000000,1000000,.1],["rope","rope_freq_scale","RoPE frequency scale","number",0,-1000,1000,.001],["rope","yarn_ext_factor","YaRN ext factor","number",-1,-1,100,.01],["rope","yarn_attn_factor","YaRN attention factor","number",1,0,100,.01],["rope","yarn_beta_fast","YaRN beta fast","number",32,0,1000,.1],["rope","yarn_beta_slow","YaRN beta slow","number",1,0,1000,.1],["rope","yarn_orig_ctx","YaRN original context","number",0,0,1000000,1]];
function q(id,label){const b=document.createElement("button");b.type="button";b.className="q";b.textContent="?";b.dataset.help=id;b.ariaLabel=label+" 설명";return b}
function field(d){const [group,id,label,type,val,a,b,c]=d,w=document.createElement("div");w.className="field"+(type==="textarea"?" all":"");if(type==="check"){w.className="check";const i=document.createElement("input");i.type="checkbox";i.id=id;i.checked=val;const l=document.createElement("label");l.htmlFor=id;l.textContent=label;w.append(i,l,q(id,label));$(group).append(w);return w}const l=document.createElement("label");l.htmlFor=id;l.append(document.createTextNode(label+" "),q(id,label));let i;if(type==="select"){i=document.createElement("select");for(const [v,t] of a){const o=document.createElement("option");o.value=v;o.textContent=t;i.append(o)}i.value=val}else if(type==="textarea"){i=document.createElement("textarea");i.value=val}else{i=document.createElement("input");i.type="number";i.value=val;i.min=a;i.max=b;i.step=c}i.id=id;w.append(l,i);$(group).append(w)}[...GEN,...RUN].forEach(field);
function openHelp(id,trigger){const x=HELP[id];if(!x)return;lastHelp=trigger;$("helpTitle").textContent=trigger?.closest("label")?.firstChild?.textContent?.trim()||id;$("helpGroup").textContent=x.group;$("helpSummary").textContent=x.summary;$("helpDefault").textContent=x.def;$("helpRange").textContent=x.range;$("helpEffect").textContent=x.effect;$("helpCaution").textContent=x.caution;$("back").hidden=false;$("close").focus()}function closeHelp(){$("back").hidden=true;lastHelp?.focus()}
document.addEventListener("click",e=>{const b=e.target.closest("[data-help]");if(b){e.preventDefault();openHelp(b.dataset.help,b)}});$("close").onclick=closeHelp;$("back").onclick=e=>{if(e.target===$("back"))closeHelp()};document.addEventListener("keydown",e=>{if(e.key==="Escape"&&!$("back").hidden)closeHelp()});
document.querySelectorAll(".tab").forEach(t=>t.onclick=()=>{document.querySelectorAll(".tab").forEach(x=>x.classList.remove("on"));t.classList.add("on");const g=t.dataset.pane==="gen";$("genPane").classList.toggle("hide",!g);$("runPane").classList.toggle("hide",g)});
function num(id,d=0){const x=Number($(id).value);return Number.isFinite(x)?x:d}function tri(id){return $(id).value==="auto"?null:$(id).value==="true"}function optInt(id){const s=$(id).value.trim();return s===""?null:Math.trunc(Number(s))}
function miro(){const on=num("mirostat_mode")!==0;$("miroNote").classList.toggle("hide",!on);["top_p","top_k","min_p","typical_p","tfs_z"].forEach(id=>$(id).disabled=on)}$("mirostat_mode").onchange=miro;miro();$("grammar_mode").onchange=()=>$("grammar_text").disabled=$("grammar_mode").value==="none";
$("precise").onclick=()=>{Object.entries({temperature:.2,top_p:.85,top_k:20,min_p:.05,typical_p:1,tfs_z:1,repeat_penalty:1.08,mirostat_mode:0}).forEach(([k,v])=>$(k).value=v);miro()};$("creative").onclick=()=>{Object.entries({temperature:1,top_p:.95,top_k:80,min_p:.02,typical_p:1,tfs_z:1,repeat_penalty:1.03,mirostat_mode:0}).forEach(([k,v])=>$(k).value=v);miro()};
function genPayload(prompt){return{system_prompt:$("system_prompt").value,history,prompt,temperature:num("temperature",.7),top_p:num("top_p",.9),top_k:Math.trunc(num("top_k",40)),min_p:num("min_p",.05),typical_p:num("typical_p",1),tfs_z:num("tfs_z",1),repeat_penalty:num("repeat_penalty",1.1),presence_penalty:num("presence_penalty"),frequency_penalty:num("frequency_penalty"),seed:optInt("seed"),max_tokens:Math.trunc(num("max_tokens",256)),top_logprobs:Math.trunc(num("top_logprobs",15)),response_format:$("response_format").value,stop:$("stop").value.split("\n").map(x=>x.trim()).filter(Boolean),mirostat_mode:Math.trunc(num("mirostat_mode")),mirostat_tau:num("mirostat_tau",5),mirostat_eta:num("mirostat_eta",.1),logit_bias:$("logit_bias").value.trim(),grammar_mode:$("grammar_mode").value,grammar_text:$("grammar_text").value}}
function runPayload(){return{n_ctx:Math.trunc(num("n_ctx",4096)),n_batch:Math.trunc(num("n_batch",512)),n_ubatch:Math.trunc(num("n_ubatch",512)),last_n_tokens_size:Math.trunc(num("last_n_tokens_size",64)),n_threads:Math.trunc(num("n_threads",2)),n_threads_batch:Math.trunc(num("n_threads_batch",4)),n_gpu_layers:Math.trunc(num("n_gpu_layers")),numa:$("numa").checked,use_mmap:$("use_mmap").checked,use_mlock:$("use_mlock").checked,offload_kqv:$("offload_kqv").checked,flash_attn:$("flash_attn").checked,op_offload:tri("op_offload"),swa_full:tri("swa_full"),rope_freq_base:num("rope_freq_base"),rope_freq_scale:num("rope_freq_scale"),yarn_ext_factor:num("yarn_ext_factor",-1),yarn_attn_factor:num("yarn_attn_factor",1),yarn_beta_fast:num("yarn_beta_fast",32),yarn_beta_slow:num("yarn_beta_slow",1),yarn_orig_ctx:Math.trunc(num("yarn_orig_ctx"))}}
function status(kind,text){$("dot").className="dot"+(kind==="load"?" load":kind==="err"?" err":"");$("status").textContent=text}async function refresh(){try{const d=await(await fetch("/api/status",{cache:"no-store"})).json();status(d.loading?"load":d.ready?"ok":"err",d.loading?"모델 로딩 중":d.ready?"모델 준비됨":d.error||"모델 없음");if(d.settings)for(const[k,v]of Object.entries(d.settings)){const i=$(k);if(!i)continue;if(i.type==="checkbox")i.checked=!!v;else if(v===null&&(k==="op_offload"||k==="swa_full"))i.value="auto";else i.value=String(v)}}catch{status("err","서버 연결 실패")}}refresh();
function add(role,text,tokens){$("empty")?.remove();const w=document.createElement("div");w.className="msg "+role;const r=document.createElement("div");r.className="role";r.textContent=role==="user"?"USER":"ASSISTANT";const b=document.createElement("div");b.className="bubble";if(tokens?.length)tokens.forEach((t,i)=>{const s=document.createElement("span");s.className="tok";s.textContent=t.token;s.onclick=()=>inspect(t,i,s);b.append(s)});else b.textContent=text;w.append(r,b);$("messages").append(w);$("messages").scrollTop=$("messages").scrollHeight}
const vis=s=>(s||"∅").replaceAll(" ","␠").replaceAll("\n","↵\n").replaceAll("\t","⇥");function inspect(t,i,el){selected?.classList.remove("sel");selected=el;el.classList.add("sel");$("inspector").classList.add("open");const b=$("inspectBody");b.innerHTML="";const title=document.createElement("div");title.className="tokenTitle";title.textContent=vis(t.token);const st=document.createElement("div");st.className="stats";st.innerHTML=`<div class=stat><span class=small>순번</span><b>#${i+1}</b></div><div class=stat><span class=small>선택 확률</span><b>${(t.probability*100).toFixed(5)}%</b></div><div class=stat><span class=small>Log probability</span><b>${t.logprob==null?"N/A":Number(t.logprob).toFixed(6)}</b></div><div class=stat><span class=small>후보 수</span><b>${t.candidates.length}</b></div>`;b.append(title,st);const mx=Math.max(...t.candidates.map(x=>x.probability||0),1e-12);for(const c of t.candidates){const w=document.createElement("div");w.className="cand";w.innerHTML=`<div class=candHead><span class="candName ${c.chosen?"chosen":""}">${vis(c.token)}${c.chosen?" ← 선택":""}</span><span>${(c.probability*100).toFixed(5)}%</span></div><div class=bar><div style="width:${Math.max(.3,c.probability/mx*100)}%"></div></div>`;b.append(w)}}
async function send(){if(busy)return;const p=$("prompt").value.trim();if(!p)return;busy=true;$("send").disabled=$("prompt").disabled=true;$("error").textContent="";status("load","생성 중");add("user",p);$("prompt").value="";try{const r=await fetch("/api/generate",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(genPayload(p))}),d=await r.json();if(!r.ok)throw Error(d.error||r.status);add("assistant",d.text,d.tokens);history.push({role:"user",content:p},{role:"assistant",content:d.text});while(history.length>20)history.shift();const u=d.usage||{};$("usage").textContent=`입력 ${u.prompt_tokens??"?"} · 출력 ${u.completion_tokens??"?"} · 전체 ${u.total_tokens??"?"} · 종료 ${d.finish_reason??"?"}`}catch(e){$("error").textContent="오류: "+e.message;add("assistant","[생성 실패]")}finally{busy=false;$("send").disabled=$("prompt").disabled=false;status("ok","모델 준비됨");$("prompt").focus()}}$("send").onclick=send;$("prompt").onkeydown=e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();send()}};
function reset(msg="대화를 초기화했습니다."){history.length=0;selected=null;$("messages").innerHTML=`<div id=empty class=empty><h2>${msg}</h2><p>새 메시지를 입력하세요.</p></div>`;$("inspectBody").textContent="답변에서 토큰을 선택하세요.";$("usage").textContent="토큰을 클릭하면 후보 확률을 볼 수 있습니다."}$("reset").onclick=()=>reset();$("reload").onclick=async()=>{if(busy)return;busy=true;$("reload").disabled=$("send").disabled=true;status("load","모델 재로딩 중");try{const r=await fetch("/api/reload",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(runPayload())}),d=await r.json();if(!r.ok)throw Error(d.error||r.status);reset("새 실행 설정으로 모델을 다시 로드했습니다.");status("ok","모델 준비됨")}catch(e){$("error").textContent="재로드 오류: "+e.message;status("err","재로드 실패")}finally{busy=false;$("reload").disabled=$("send").disabled=false}};
</script>
</body>
</html>'''


def clamp_float(value: Any, low: float, high: float, default: float) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return default
    return max(low, min(high, number)) if math.isfinite(number) else default


def clamp_int(value: Any, low: int, high: int, default: int) -> int:
    try:
        return max(low, min(high, int(value)))
    except (TypeError, ValueError):
        return default


def tri_bool(value: Any) -> bool | None:
    if value is None or isinstance(value, bool):
        return value
    raise ValueError("삼중 상태 값은 true, false 또는 null이어야 합니다.")


def logprob_number(value: Any) -> float | None:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def probability(value: Any) -> float:
    number = logprob_number(value)
    return 0.0 if number is None else math.exp(max(-745.0, min(0.0, number)))


def clean_history(value: Any) -> list[dict[str, str]]:
    result: list[dict[str, str]] = []
    if not isinstance(value, list):
        return result
    for item in value[-20:]:
        if not isinstance(item, dict):
            continue
        role, content = item.get("role"), item.get("content")
        if role in {"user", "assistant"} and isinstance(content, str):
            result.append({"role": role, "content": content[:12_000]})
    return result


def parse_bias(raw: Any) -> dict[int, float] | None:
    if raw in (None, ""):
        return None
    try:
        data = json.loads(raw) if isinstance(raw, str) else raw
    except json.JSONDecodeError as error:
        raise ValueError(f"logit_bias JSON 오류: {error}") from error
    if not isinstance(data, dict):
        raise ValueError("logit_bias는 JSON 객체여야 합니다.")
    result: dict[int, float] = {}
    for key, value in data.items():
        try:
            result[int(key)] = clamp_float(value, -100.0, 100.0, 0.0)
        except (TypeError, ValueError) as error:
            raise ValueError("logit_bias 키는 token ID여야 합니다.") from error
    return result or None


def make_grammar(mode: Any, text: Any) -> Any:
    if mode in (None, "", "none"):
        return None
    if not isinstance(text, str) or not text.strip():
        raise ValueError("Grammar 내용을 입력하십시오.")
    if mode == "gbnf":
        return GRAMMAR_CLASS.from_string(text, verbose=False)
    if mode == "json_schema":
        json.loads(text)
        return GRAMMAR_CLASS.from_json_schema(text, verbose=False)
    raise ValueError("지원하지 않는 Grammar mode입니다.")


def token_details(choice: dict[str, Any]) -> list[dict[str, Any]]:
    content = (choice.get("logprobs") or {}).get("content")
    if not isinstance(content, list):
        return []
    result = []
    for item in content:
        if not isinstance(item, dict):
            continue
        selected = str(item.get("token", ""))
        selected_lp = logprob_number(item.get("logprob"))
        candidates, seen = [], set()
        for candidate in item.get("top_logprobs") or []:
            if not isinstance(candidate, dict):
                continue
            token = str(candidate.get("token", ""))
            lp = logprob_number(candidate.get("logprob"))
            key = (token, lp)
            if key in seen:
                continue
            seen.add(key)
            candidates.append({
                "token": token,
                "logprob": lp,
                "probability": probability(lp),
                "chosen": token == selected,
            })
        if not any(x["chosen"] for x in candidates):
            candidates.append({
                "token": selected,
                "logprob": selected_lp,
                "probability": probability(selected_lp),
                "chosen": True,
            })
        candidates.sort(key=lambda x: x["probability"], reverse=True)
        result.append({
            "token": selected,
            "logprob": selected_lp,
            "probability": probability(selected_lp),
            "candidates": candidates,
        })
    return result


def runtime_values(data: dict[str, Any]) -> dict[str, Any]:
    n_ctx = clamp_int(data.get("n_ctx"), 256, 32768, 4096)
    n_batch = clamp_int(data.get("n_batch"), 1, n_ctx, min(512, n_ctx))
    return {
        "n_ctx": n_ctx,
        "n_batch": n_batch,
        "n_ubatch": clamp_int(
            data.get("n_ubatch"),
            1,
            n_batch,
            min(512, n_batch),
        ),
        "n_threads": clamp_int(data.get("n_threads"), 1, 256, 2),
        "n_threads_batch": clamp_int(
            data.get("n_threads_batch"),
            1,
            256,
            4,
        ),
        "n_gpu_layers": clamp_int(
            data.get("n_gpu_layers"),
            -1,
            200,
            0,
        ),
        "last_n_tokens_size": clamp_int(
            data.get("last_n_tokens_size"),
            0,
            n_ctx,
            64,
        ),
        "use_mmap": bool(data.get("use_mmap", True)),
        "use_mlock": bool(data.get("use_mlock", False)),
        "offload_kqv": bool(data.get("offload_kqv", True)),
        "flash_attn": bool(data.get("flash_attn", False)),
        "op_offload": tri_bool(data.get("op_offload")),
        "swa_full": tri_bool(data.get("swa_full")),
        "numa": bool(data.get("numa", False)),
        "rope_freq_base": clamp_float(
            data.get("rope_freq_base"),
            -1e6,
            1e6,
            0.0,
        ),
        "rope_freq_scale": clamp_float(
            data.get("rope_freq_scale"),
            -1000,
            1000,
            0.0,
        ),
        "yarn_ext_factor": clamp_float(
            data.get("yarn_ext_factor"),
            -1,
            100,
            -1.0,
        ),
        "yarn_attn_factor": clamp_float(
            data.get("yarn_attn_factor"),
            0,
            100,
            1.0,
        ),
        "yarn_beta_fast": clamp_float(
            data.get("yarn_beta_fast"),
            0,
            1000,
            32.0,
        ),
        "yarn_beta_slow": clamp_float(
            data.get("yarn_beta_slow"),
            0,
            1000,
            1.0,
        ),
        "yarn_orig_ctx": clamp_int(
            data.get("yarn_orig_ctx"),
            0,
            1_000_000,
            0,
        ),
    }


def build_model(settings: dict[str, Any]) -> Any:
    from llama_cpp import Llama

    kwargs = {
        "repo_id": MODEL_REPO,
        "filename": MODEL_FILE,
        **settings,
        "logits_all": True,
        "verbose": False,
    }

    if settings["op_offload"] is None:
        kwargs.pop("op_offload")

    if settings["swa_full"] is None:
        kwargs.pop("swa_full")

    print(
        f"[MODEL] loading {MODEL_FILE}, "
        f"n_ctx={settings['n_ctx']}, "
        f"gpu_layers={settings['n_gpu_layers']}"
    )

    return Llama.from_pretrained(**kwargs)


def reload_model(settings: dict[str, Any]) -> None:
    global MODEL, RUNTIME

    with MODEL_LOCK:
        STATUS.update(
            ready=False,
            loading=True,
            error=None,
        )

        old = MODEL
        MODEL = None

        if old is not None:
            del old
            gc.collect()

        try:
            MODEL = build_model(settings)
            RUNTIME = dict(settings)

            STATUS.update(
                ready=True,
                loading=False,
                error=None,
            )

        except Exception as error:
            STATUS.update(
                ready=False,
                loading=False,
                error=f"{type(error).__name__}: {error}",
            )
            raise


def generate(data: dict[str, Any]) -> dict[str, Any]:
    if MODEL is None or not STATUS["ready"]:
        raise RuntimeError("모델이 준비되지 않았습니다.")

    prompt = data.get("prompt")

    if not isinstance(prompt, str) or not prompt.strip():
        raise ValueError("prompt가 비어 있습니다.")

    messages = []
    system = data.get("system_prompt")

    if isinstance(system, str) and system.strip():
        messages.append({
            "role": "system",
            "content": system[:12_000],
        })

    messages.extend(clean_history(data.get("history")))

    messages.append({
        "role": "user",
        "content": prompt[:12_000],
    })

    stop = data.get("stop")

    stop = (
        [
            str(item)[:200]
            for item in stop[:20]
            if str(item)
        ]
        if isinstance(stop, list)
        else []
    )

    response_format = (
        {"type": "json_object"}
        if data.get("response_format") == "json_object"
        else None
    )

    grammar = make_grammar(
        data.get("grammar_mode"),
        data.get("grammar_text"),
    )

    if grammar is not None and response_format is not None:
        raise ValueError(
            "Grammar와 JSON response format을 동시에 사용하지 마십시오."
        )

    seed_value = data.get("seed")

    seed = (
        None
        if seed_value is None
        else clamp_int(seed_value, 0, 2**31 - 1, 0)
    )

    with MODEL_LOCK:
        response = MODEL.create_chat_completion(
            messages=messages,
            temperature=clamp_float(
                data.get("temperature"),
                -1,
                2,
                0.7,
            ),
            top_p=clamp_float(
                data.get("top_p"),
                0,
                1,
                0.9,
            ),
            top_k=clamp_int(
                data.get("top_k"),
                0,
                100_000,
                40,
            ),
            min_p=clamp_float(
                data.get("min_p"),
                0,
                1,
                0.05,
            ),
            typical_p=clamp_float(
                data.get("typical_p"),
                0,
                1,
                1.0,
            ),
            tfs_z=clamp_float(
                data.get("tfs_z"),
                0,
                2,
                1.0,
            ),
            repeat_penalty=clamp_float(
                data.get("repeat_penalty"),
                0,
                3,
                1.1,
            ),
            presence_penalty=clamp_float(
                data.get("presence_penalty"),
                -2,
                2,
                0.0,
            ),
            frequency_penalty=clamp_float(
                data.get("frequency_penalty"),
                -2,
                2,
                0.0,
            ),
            max_tokens=clamp_int(
                data.get("max_tokens"),
                1,
                min(4096, RUNTIME["n_ctx"]),
                256,
            ),
            seed=seed,
            stop=stop,
            response_format=response_format,
            mirostat_mode=clamp_int(
                data.get("mirostat_mode"),
                0,
                2,
                0,
            ),
            mirostat_tau=clamp_float(
                data.get("mirostat_tau"),
                0,
                20,
                5.0,
            ),
            mirostat_eta=clamp_float(
                data.get("mirostat_eta"),
                0,
                1,
                0.1,
            ),
            grammar=grammar,
            logit_bias=parse_bias(
                data.get("logit_bias")
            ),
            logprobs=True,
            top_logprobs=clamp_int(
                data.get("top_logprobs"),
                1,
                50,
                15,
            ),
            stream=False,
        )

    choice = response["choices"][0]

    text = (
        (choice.get("message") or {}).get("content")
        or ""
    )

    return {
        "text": str(text),
        "tokens": token_details(choice),
        "finish_reason": choice.get("finish_reason"),
        "usage": response.get("usage", {}),
    }


class Handler(BaseHTTPRequestHandler):
    server_version = "QwenHyperparameterLab/4.0"

    def log_message(
        self,
        fmt: str,
        *args: Any,
    ) -> None:
        print(
            f"[HTTP] {self.address_string()} - "
            f"{fmt % args}"
        )

    def send_bytes(
        self,
        status: int,
        content_type: str,
        body: bytes,
    ) -> None:
        self.send_response(status)
        self.send_header(
            "Content-Type",
            content_type,
        )
        self.send_header(
            "Content-Length",
            str(len(body)),
        )
        self.send_header(
            "Cache-Control",
            "no-store",
        )
        self.send_header(
            "X-Content-Type-Options",
            "nosniff",
        )
        self.end_headers()
        self.wfile.write(body)

    def send_json(
        self,
        status: int,
        data: dict[str, Any],
    ) -> None:
        body = json.dumps(
            data,
            ensure_ascii=False,
            allow_nan=False,
        ).encode("utf-8")

        self.send_bytes(
            status,
            "application/json; charset=utf-8",
            body,
        )

    def read_json(self) -> dict[str, Any]:
        length = int(
            self.headers.get(
                "Content-Length",
                "0",
            )
        )

        if length <= 0 or length > 3_000_000:
            raise ValueError(
                "요청 본문 크기가 올바르지 않습니다."
            )

        data = json.loads(
            self.rfile.read(length).decode("utf-8")
        )

        if not isinstance(data, dict):
            raise ValueError(
                "JSON 객체가 필요합니다."
            )

        return data

    def do_GET(self) -> None:
        path = urlparse(self.path).path

        if path == "/":
            self.send_bytes(
                200,
                "text/html; charset=utf-8",
                HTML.encode("utf-8"),
            )

        elif path == "/api/status":
            self.send_json(
                200,
                {
                    **STATUS,
                    "settings": RUNTIME,
                },
            )

        elif path == "/favicon.ico":
            self.send_bytes(
                204,
                "image/x-icon",
                b"",
            )

        else:
            self.send_json(
                404,
                {
                    "error": "찾을 수 없는 경로입니다."
                },
            )

    def do_POST(self) -> None:
        try:
            path = urlparse(self.path).path
            data = self.read_json()

            if path == "/api/generate":
                self.send_json(
                    200,
                    generate(data),
                )

            elif path == "/api/reload":
                reload_model(
                    runtime_values(data)
                )

                self.send_json(
                    200,
                    {
                        "ok": True,
                        "settings": RUNTIME,
                    },
                )

            else:
                self.send_json(
                    404,
                    {
                        "error": "찾을 수 없는 경로입니다."
                    },
                )

        except (
            ValueError,
            json.JSONDecodeError,
        ) as error:
            self.send_json(
                400,
                {
                    "error": str(error)
                },
            )

        except Exception as error:
            print(
                f"[ERROR] "
                f"{type(error).__name__}: "
                f"{error}"
            )

            self.send_json(
                500,
                {
                    "error": (
                        f"{type(error).__name__}: "
                        f"{error}"
                    )
                },
            )


def main() -> None:
    global GRAMMAR_CLASS

    try:
        from llama_cpp import LlamaGrammar

        GRAMMAR_CLASS = LlamaGrammar

    except ImportError as error:
        print(
            "llama-cpp-python이 설치되지 않았습니다."
        )
        raise SystemExit(1) from error

    reload_model(DEFAULT_RUNTIME)

    server = ThreadingHTTPServer(
        (HOST, PORT),
        Handler,
    )

    server.daemon_threads = True

    url = f"http://{HOST}:{PORT}"

    print(
        f"[SERVER] {url}\n"
        "[SERVER] 종료: Ctrl+C"
    )

    timer = threading.Timer(
        0.8,
        lambda: webbrowser.open(url),
    )
    timer.daemon = True
    timer.start()

    try:
        server.serve_forever()

    except KeyboardInterrupt:
        print("\n[SERVER] 종료합니다.")

    finally:
        server.server_close()


if __name__ == "__main__":
    main()