# pptx2markdown 에이전트 설치 지침

[English](agent-install.md) | 한국어

당신은 사용자의 컴퓨터에 `pptx2markdown`을 설치하는 AI 코딩 에이전트입니다.
사용자가 개발용 설치를 명시적으로 요청하지 않았다면 저장소 clone, 로컬 wheel,
`dist/`, 다른 로컬 빌드 산출물이 아니라 PyPI에 공개된 패키지를 설치하세요.

명령 실행, Python 또는 격리된 Python 도구 관리자 설치, 셸 profile 수정,
LibreOffice 설치에 권한이 필요할 수 있습니다. 상태를 바꾸는 명령은 실행 전에
각각 설명하세요. 관리자 권한, `sudo`, 비밀번호, 시스템 패키지 관리자, 셸
profile, PATH 변경, LibreOffice 설치 전에는 반드시 사용자에게 확인하세요.

## 1. 대화 언어 선택

설치 안내와 최종 보고에 사용할 언어를 먼저 확인하세요. 설치 과정 전체에서 그
언어를 사용하세요.

## 2. 환경 확인

읽기 전용 명령으로 운영체제, CPU 아키텍처, 현재 셸, 사용 가능한 Python 및 도구
관리자를 확인하세요. 해당되는 경우 다음을 점검합니다.

```bash
python3 --version
python --version
uv --version
pipx --version
```

`pptx2markdown`에는 Python 3.12 이상이 필요합니다. 시스템 Python을 교체하거나
기본 interpreter를 자동으로 바꾸지 마세요. 적합한 interpreter가 없다면 플랫폼에
맞는 공식 설치 선택지를 설명하고, 무엇을 사용할지 확인한 후 설치하세요.

에이전트가 명령을 실행할 수 없다면 작업을 중단하고 README의 수동 설치 절차를
안내하세요. 설치가 성공했다고 말하면 안 됩니다.

## 3. 설치 범위 확인

다음 선택지를 제시하세요.

1. `uv tool`을 이용한 격리 CLI 설치(`uv`가 있으면 권장)
2. `pipx`를 이용한 격리 CLI 설치
3. 현재 활성화된 가상환경에 설치

관리자 권한으로 시스템 Python에 설치하지 마세요. 경로를 확인하지 않은 임의의
디렉터리에 새 가상환경을 만들지 마세요.

## 4. PyPI에서 설치

사용자가 선택한 방법에 맞는 명령을 사용하세요.

`uv`:

```bash
uv tool install --upgrade pptx2markdown
```

`pipx`는 패키지 설치 여부를 확인한 다음 설치하거나 업그레이드합니다.

```bash
pipx install pptx2markdown
pipx upgrade pptx2markdown
```

사용자가 승인한 활성 가상환경:

```bash
python -m pip install --upgrade pptx2markdown
```

명령 디렉터리가 PATH에 없다면 도구가 출력한 정확한 안내를 보고하세요.
`uv tool update-shell`, `pipx ensurepath` 같은 셸 profile 변경 명령을 실행하기 전에
반드시 확인하세요. 가능하면 profile을 몰래 수정하지 말고 셸을 새로 열거나
새로고침하는 방법을 안내하세요.

## 5. 구형 PPT 지원 필요 여부 확인

일반 `.pptx` 변환에는 LibreOffice가 필요하지 않습니다. LibreOffice는 구형 `.ppt`
입력과 EMF/WMF 처리에 선택적으로 사용됩니다.

사용자에게 이 기능이 필요한지 물으세요. 필요하지 않으면 LibreOffice를 설치하지
마세요. 필요하다면 `soffice` 또는 LibreOffice가 이미 있는지 먼저 확인하세요.
Homebrew, APT, 다른 시스템 패키지 관리자, 관리자 권한, `sudo`를 사용하기 전에
정확한 명령을 설명하고 승인을 받으세요.

Microsoft PowerPoint를 설치하거나 Office 설정을 바꾸지 마세요.

## 6. 설치 검증

PATH 변경으로 필요하다면 새 셸 또는 새로고침된 셸에서 다음을 실행하세요.

```bash
pptx2markdown --help
```

도움말에 `--output-format {markdown,json}`와 `--headings {auto,strict}`를 포함한
정적 파서 옵션이 표시되는지 확인하세요. Surya, OCR, VLM, 읽기 순서 provider
옵션이 표시되면 안 됩니다.

설치 방법이 패키지 목록 기능을 제공하면 `uv tool list`, `pipx list`, Python
package metadata 등으로 설치된 버전을 보고하세요.

설치 확인만을 위해 사용자의 개인 프레젠테이션을 변환하지 마세요. 사용자가 파일을
제공하거나 선택하고 출력 디렉터리를 승인한 경우에만 변환 테스트를 제안하세요.

## 7. 최종 보고

검증한 사실만 요약하세요.

- 확인한 운영체제, 셸, Python 버전
- 선택한 설치 방식과 설치된 패키지 버전
- CLI 도움말 검증 성공 여부
- LibreOffice 설치, 기존 설치 확인, 건너뜀, 거절 여부
- 필요한 PATH 새로고침 또는 남은 수동 작업
- 첫 변환에 사용할 정확한 명령

```bash
pptx2markdown deck.pptx
```

기본 출력이 `./output/<deck-name>/` 아래에 생성된다고 설명하세요. 실패한 단계가
있으면 설치가 완료됐다고 말하지 말고 실패 원인과 가장 안전한 다음 조치를
보고하세요.
