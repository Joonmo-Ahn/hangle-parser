"""
비동기 프로그래밍 예제 - asyncio + httpx

이 모듈은 동기/비동기 방식의 성능 차이를 비교하고,
asyncio의 핵심 개념들을 실제 사용 사례와 함께 설명합니다.

핵심 개념:
    1. 코루틴(Coroutine): async def로 정의된 비동기 함수
    2. 이벤트 루프(Event Loop): 코루틴의 실행을 관리하는 스케줄러
    3. Task: 코루틴을 이벤트 루프에서 실행 가능한 단위로 감싼 것
    4. await: 비동기 작업의 완료를 기다리는 키워드

왜 비동기가 I/O 바운드 작업에서 우수한가?
    - 동기 방식: 네트워크 응답을 기다리는 동안 CPU가 유휴 상태
    - 비동기 방식: 응답 대기 중 다른 작업을 처리하여 CPU 활용도 극대화
    - N개의 요청이 각각 T초 걸릴 때:
        * 동기: N × T초 (순차 실행)
        * 비동기: ~T초 (동시 실행, 이론상)

실행 방법:
    pip install httpx
    python async_demo.py

작성자: Claude
버전: 1.0.0
"""

from __future__ import annotations

import asyncio
import time
from dataclasses import dataclass, field
from typing import Any, Callable, Coroutine
from enum import Enum
from contextlib import contextmanager

# httpx는 동기/비동기 모두 지원하는 현대적인 HTTP 클라이언트
try:
    import httpx
    HAS_HTTPX = True
except ImportError:
    HAS_HTTPX = False
    print("⚠️  httpx 라이브러리가 필요합니다: pip install httpx")


# =============================================================================
# 상수 및 설정
# =============================================================================

class APIEndpoint(str, Enum):
    """테스트용 공개 API 엔드포인트"""
    POSTS = "https://jsonplaceholder.typicode.com/posts"
    USERS = "https://jsonplaceholder.typicode.com/users"
    COMMENTS = "https://jsonplaceholder.typicode.com/comments"
    TODOS = "https://jsonplaceholder.typicode.com/todos"
    ALBUMS = "https://jsonplaceholder.typicode.com/albums"


# 테스트할 URL 목록 (다양한 엔드포인트)
TEST_URLS: list[str] = [
    f"{APIEndpoint.POSTS.value}/{i}" for i in range(1, 11)  # 10개 포스트
] + [
    f"{APIEndpoint.USERS.value}/{i}" for i in range(1, 6)   # 5개 유저
] + [
    f"{APIEndpoint.TODOS.value}/{i}" for i in range(1, 6)   # 5개 할일
]

# 요청 제한 설정 (서버 부하 방지)
MAX_CONCURRENT_REQUESTS: int = 10
REQUEST_TIMEOUT_SECONDS: float = 10.0


# =============================================================================
# 데이터 클래스
# =============================================================================

@dataclass
class FetchResult:
    """
    HTTP 요청 결과를 담는 데이터 클래스

    Attributes:
        url: 요청한 URL
        status_code: HTTP 상태 코드
        data: 응답 데이터 (JSON 파싱된 결과)
        elapsed_ms: 요청 소요 시간 (밀리초)
        success: 성공 여부
        error: 에러 메시지 (실패 시)
    """
    url: str
    status_code: int = 0
    data: dict[str, Any] | list[Any] | None = None
    elapsed_ms: float = 0.0
    success: bool = False
    error: str | None = None

    def __repr__(self) -> str:
        status = "✓" if self.success else "✗"
        return f"[{status}] {self.url} - {self.status_code} ({self.elapsed_ms:.1f}ms)"


@dataclass
class BenchmarkResult:
    """
    성능 벤치마크 결과

    Attributes:
        method: 실행 방식 (sync/async)
        total_requests: 총 요청 수
        successful: 성공한 요청 수
        failed: 실패한 요청 수
        total_time_sec: 전체 소요 시간 (초)
        avg_request_ms: 평균 요청 시간 (밀리초)
        requests_per_sec: 초당 처리 요청 수
    """
    method: str
    total_requests: int = 0
    successful: int = 0
    failed: int = 0
    total_time_sec: float = 0.0
    avg_request_ms: float = 0.0
    requests_per_sec: float = 0.0
    results: list[FetchResult] = field(default_factory=list)

    def calculate_stats(self) -> None:
        """통계 계산"""
        if self.results:
            self.total_requests = len(self.results)
            self.successful = sum(1 for r in self.results if r.success)
            self.failed = self.total_requests - self.successful

            successful_times = [r.elapsed_ms for r in self.results if r.success]
            if successful_times:
                self.avg_request_ms = sum(successful_times) / len(successful_times)

            if self.total_time_sec > 0:
                self.requests_per_sec = self.total_requests / self.total_time_sec


# =============================================================================
# 유틸리티 함수
# =============================================================================

@contextmanager
def timer(description: str = ""):
    """
    실행 시간을 측정하는 컨텍스트 매니저

    사용 예시:
        with timer("작업 A") as t:
            do_something()
        print(f"소요 시간: {t.elapsed}초")
    """
    class TimerContext:
        elapsed: float = 0.0

    ctx = TimerContext()
    start = time.perf_counter()

    try:
        yield ctx
    finally:
        ctx.elapsed = time.perf_counter() - start
        if description:
            print(f"⏱️  {description}: {ctx.elapsed:.3f}초")


def print_section(title: str, width: int = 70) -> None:
    """섹션 구분선 출력"""
    print("\n" + "=" * width)
    print(f" {title}")
    print("=" * width)


def print_benchmark_comparison(sync_result: BenchmarkResult, async_result: BenchmarkResult) -> None:
    """
    동기/비동기 벤치마크 결과 비교 출력

    Args:
        sync_result: 동기 방식 결과
        async_result: 비동기 방식 결과
    """
    print_section("📊 성능 비교 분석 결과")

    # 표 형식 출력
    print(f"\n{'항목':<25} {'동기(Sync)':<20} {'비동기(Async)':<20}")
    print("-" * 65)
    print(f"{'총 요청 수':<25} {sync_result.total_requests:<20} {async_result.total_requests:<20}")
    print(f"{'성공':<25} {sync_result.successful:<20} {async_result.successful:<20}")
    print(f"{'실패':<25} {sync_result.failed:<20} {async_result.failed:<20}")
    print(f"{'총 소요 시간':<25} {sync_result.total_time_sec:<20.3f} {async_result.total_time_sec:<20.3f}")
    print(f"{'평균 요청 시간(ms)':<25} {sync_result.avg_request_ms:<20.1f} {async_result.avg_request_ms:<20.1f}")
    print(f"{'초당 처리량':<25} {sync_result.requests_per_sec:<20.1f} {async_result.requests_per_sec:<20.1f}")

    # 성능 향상 비율 계산
    if async_result.total_time_sec > 0:
        speedup = sync_result.total_time_sec / async_result.total_time_sec
        print(f"\n🚀 비동기 방식이 약 {speedup:.1f}배 빠릅니다!")

    # 분석 설명
    print("\n📝 분석:")
    print("   • 동기 방식: 각 요청이 완료될 때까지 다음 요청을 보내지 않음")
    print("   • 비동기 방식: 응답을 기다리는 동안 다른 요청을 동시에 처리")
    print("   • I/O 대기 시간이 많을수록 비동기의 이점이 크게 나타남")
    print(f"   • 이론적 최대 속도 향상: {sync_result.total_requests}배 (실제는 네트워크/서버 제한)")


# =============================================================================
# 동기 방식 구현
# =============================================================================

def fetch_sync(url: str, client: httpx.Client) -> FetchResult:
    """
    동기 방식 HTTP GET 요청

    Args:
        url: 요청할 URL
        client: httpx 동기 클라이언트

    Returns:
        FetchResult: 요청 결과
    """
    result = FetchResult(url=url)
    start = time.perf_counter()

    try:
        response = client.get(url, timeout=REQUEST_TIMEOUT_SECONDS)
        result.status_code = response.status_code
        result.data = response.json()
        result.success = response.is_success

    except httpx.TimeoutException:
        result.error = "타임아웃 발생"
    except httpx.RequestError as e:
        result.error = f"요청 오류: {type(e).__name__}"
    except Exception as e:
        result.error = f"알 수 없는 오류: {str(e)}"

    result.elapsed_ms = (time.perf_counter() - start) * 1000
    return result


def run_sync_benchmark(urls: list[str]) -> BenchmarkResult:
    """
    동기 방식으로 모든 URL 요청 실행

    동작 방식:
        1. 첫 번째 URL 요청 → 응답 대기 → 완료
        2. 두 번째 URL 요청 → 응답 대기 → 완료
        3. ... (순차적으로 반복)

    문제점:
        - 네트워크 I/O 대기 시간 동안 CPU가 아무것도 하지 않음
        - 총 시간 = 개별 요청 시간의 합

    Args:
        urls: 요청할 URL 목록

    Returns:
        BenchmarkResult: 벤치마크 결과
    """
    print_section("🐢 동기(Synchronous) 방식 실행")
    print(f"총 {len(urls)}개의 URL을 순차적으로 요청합니다...")

    benchmark = BenchmarkResult(method="sync")

    with timer("동기 방식 전체 소요 시간") as t:
        # 커넥션 풀을 재사용하기 위해 Client 컨텍스트 사용
        with httpx.Client() as client:
            for i, url in enumerate(urls, 1):
                result = fetch_sync(url, client)
                benchmark.results.append(result)

                # 진행 상황 출력 (10개마다)
                if i % 10 == 0 or i == len(urls):
                    success_count = sum(1 for r in benchmark.results if r.success)
                    print(f"   진행: {i}/{len(urls)} (성공: {success_count})")

    benchmark.total_time_sec = t.elapsed
    benchmark.calculate_stats()

    return benchmark


# =============================================================================
# 비동기 방식 구현
# =============================================================================

async def fetch_async(
    url: str,
    client: httpx.AsyncClient,
    semaphore: asyncio.Semaphore
) -> FetchResult:
    """
    비동기 방식 HTTP GET 요청

    핵심 개념:
        - async def: 이 함수가 코루틴임을 선언
        - await: 비동기 작업이 완료될 때까지 "양보"
        - Semaphore: 동시 요청 수를 제한하여 서버 부하 방지

    Args:
        url: 요청할 URL
        client: httpx 비동기 클라이언트
        semaphore: 동시성 제어용 세마포어

    Returns:
        FetchResult: 요청 결과
    """
    result = FetchResult(url=url)
    start = time.perf_counter()

    # 세마포어로 동시 요청 수 제한
    # async with: 비동기 컨텍스트 매니저
    async with semaphore:
        try:
            # await: 응답이 올 때까지 이벤트 루프에 제어권 반환
            # 이 시간 동안 다른 코루틴이 실행될 수 있음
            response = await client.get(url, timeout=REQUEST_TIMEOUT_SECONDS)
            result.status_code = response.status_code
            result.data = response.json()
            result.success = response.is_success

        except httpx.TimeoutException:
            result.error = "타임아웃 발생"
        except httpx.RequestError as e:
            result.error = f"요청 오류: {type(e).__name__}"
        except Exception as e:
            result.error = f"알 수 없는 오류: {str(e)}"

    result.elapsed_ms = (time.perf_counter() - start) * 1000
    return result


async def run_async_benchmark(urls: list[str]) -> BenchmarkResult:
    """
    비동기 방식으로 모든 URL 요청 실행

    동작 방식 (asyncio.gather 사용):
        1. 모든 URL에 대한 코루틴 객체 생성 (아직 실행 안됨)
        2. gather()로 모든 코루틴을 동시에 스케줄링
        3. 이벤트 루프가 I/O 대기 시간을 활용해 병렬 처리
        4. 모든 결과가 준비되면 반환

    이점:
        - CPU가 I/O 대기 시간 동안 다른 요청 처리
        - 총 시간 ≈ 가장 긴 개별 요청 시간 (이론상)

    Args:
        urls: 요청할 URL 목록

    Returns:
        BenchmarkResult: 벤치마크 결과
    """
    print_section("🚀 비동기(Asynchronous) 방식 실행")
    print(f"총 {len(urls)}개의 URL을 동시에 요청합니다...")
    print(f"최대 동시 요청 수: {MAX_CONCURRENT_REQUESTS}")

    benchmark = BenchmarkResult(method="async")

    # 세마포어: 동시 실행 코루틴 수 제한 (서버 과부하 방지)
    semaphore = asyncio.Semaphore(MAX_CONCURRENT_REQUESTS)

    with timer("비동기 방식 전체 소요 시간") as t:
        # 비동기 HTTP 클라이언트
        async with httpx.AsyncClient() as client:
            # 코루틴 객체 리스트 생성 (아직 실행되지 않음)
            tasks: list[Coroutine[Any, Any, FetchResult]] = [
                fetch_async(url, client, semaphore)
                for url in urls
            ]

            # asyncio.gather(): 모든 코루틴을 동시에 실행
            # return_exceptions=True: 예외 발생 시에도 다른 태스크 계속 실행
            results = await asyncio.gather(*tasks, return_exceptions=True)

            # 결과 처리
            for result in results:
                if isinstance(result, Exception):
                    # 예외가 발생한 경우
                    error_result = FetchResult(
                        url="unknown",
                        error=str(result)
                    )
                    benchmark.results.append(error_result)
                else:
                    benchmark.results.append(result)

    benchmark.total_time_sec = t.elapsed
    benchmark.calculate_stats()

    # 결과 요약 출력
    print(f"   완료: {benchmark.successful}/{benchmark.total_requests} 성공")

    return benchmark


# =============================================================================
# 고급 예제: 에러 처리와 재시도 로직
# =============================================================================

async def fetch_with_retry(
    url: str,
    client: httpx.AsyncClient,
    max_retries: int = 3,
    backoff_factor: float = 0.5
) -> FetchResult:
    """
    재시도 로직이 포함된 비동기 HTTP 요청

    지수 백오프(Exponential Backoff):
        - 실패 시 대기 시간을 점점 늘려가며 재시도
        - 서버 부하 분산 및 일시적 오류 복구에 효과적

    Args:
        url: 요청할 URL
        client: httpx 비동기 클라이언트
        max_retries: 최대 재시도 횟수
        backoff_factor: 백오프 배수 (대기시간 = backoff_factor * 2^retry)

    Returns:
        FetchResult: 요청 결과
    """
    result = FetchResult(url=url)
    last_error: str | None = None

    for attempt in range(max_retries + 1):
        start = time.perf_counter()

        try:
            response = await client.get(url, timeout=REQUEST_TIMEOUT_SECONDS)
            result.status_code = response.status_code
            result.data = response.json()
            result.success = response.is_success
            result.elapsed_ms = (time.perf_counter() - start) * 1000
            return result

        except (httpx.TimeoutException, httpx.RequestError) as e:
            last_error = f"{type(e).__name__}: {str(e)}"

            if attempt < max_retries:
                # 지수 백오프 대기
                wait_time = backoff_factor * (2 ** attempt)
                await asyncio.sleep(wait_time)

    # 모든 재시도 실패
    result.error = f"최대 재시도({max_retries}회) 초과. 마지막 오류: {last_error}"
    return result


async def process_with_progress(
    urls: list[str],
    processor: Callable[[str, httpx.AsyncClient], Coroutine[Any, Any, FetchResult]]
) -> list[FetchResult]:
    """
    진행 상황을 표시하며 비동기 요청 처리

    asyncio.as_completed() 사용:
        - 완료되는 순서대로 결과를 받을 수 있음
        - 실시간 진행 상황 표시에 유용

    Args:
        urls: 요청할 URL 목록
        processor: URL을 처리할 비동기 함수

    Returns:
        list[FetchResult]: 완료 순서대로 정렬된 결과 목록
    """
    results: list[FetchResult] = []

    async with httpx.AsyncClient() as client:
        # 태스크 생성
        tasks = {
            asyncio.create_task(processor(url, client)): url
            for url in urls
        }

        # as_completed: 완료되는 순서대로 이터레이션
        completed = 0
        for coro in asyncio.as_completed(tasks.keys()):
            result = await coro
            results.append(result)
            completed += 1

            # 진행률 표시
            progress = completed / len(urls) * 100
            status = "✓" if result.success else "✗"
            print(f"\r   [{progress:5.1f}%] {status} {result.url[-30:]:<30}", end="")

        print()  # 줄바꿈

    return results


# =============================================================================
# 고급 예제: asyncio.TaskGroup (Python 3.11+)
# =============================================================================

async def run_with_taskgroup(urls: list[str]) -> list[FetchResult]:
    """
    TaskGroup을 사용한 구조화된 동시성 (Python 3.11+)

    TaskGroup의 이점:
        - 모든 태스크가 완료될 때까지 자동 대기
        - 하나의 태스크가 실패하면 나머지 자동 취소
        - 예외가 ExceptionGroup으로 수집되어 일괄 처리 가능

    Args:
        urls: 요청할 URL 목록

    Returns:
        list[FetchResult]: 결과 목록
    """
    results: list[FetchResult] = []
    semaphore = asyncio.Semaphore(MAX_CONCURRENT_REQUESTS)

    async with httpx.AsyncClient() as client:
        try:
            # Python 3.11+ TaskGroup
            async with asyncio.TaskGroup() as tg:
                async def fetch_and_store(url: str) -> None:
                    result = await fetch_async(url, client, semaphore)
                    results.append(result)

                for url in urls:
                    tg.create_task(fetch_and_store(url))

        except* Exception as eg:
            # ExceptionGroup 처리 (Python 3.11+)
            print(f"⚠️  {len(eg.exceptions)}개의 예외 발생:")
            for exc in eg.exceptions[:3]:
                print(f"   - {type(exc).__name__}: {exc}")

    return results


# =============================================================================
# 메인 실행
# =============================================================================

async def main() -> None:
    """
    메인 비동기 함수

    실행 흐름:
        1. 동기 방식으로 모든 URL 순차 요청
        2. 비동기 방식으로 모든 URL 동시 요청
        3. 두 방식의 성능 비교 분석
    """
    print_section("🔬 Python 비동기 프로그래밍 벤치마크")
    print(f"테스트 URL 수: {len(TEST_URLS)}")
    print(f"최대 동시 요청: {MAX_CONCURRENT_REQUESTS}")
    print(f"요청 타임아웃: {REQUEST_TIMEOUT_SECONDS}초")

    # 1. 동기 방식 실행
    sync_result = run_sync_benchmark(TEST_URLS)

    # 잠시 대기 (서버 부하 분산)
    print("\n⏳ 5초 대기 후 비동기 테스트 시작...")
    await asyncio.sleep(5)

    # 2. 비동기 방식 실행
    async_result = await run_async_benchmark(TEST_URLS)

    # 3. 결과 비교
    print_benchmark_comparison(sync_result, async_result)

    # 4. 상세 결과 출력 (선택적)
    print_section("📋 개별 요청 결과 (처음 5개)")
    print("\n[동기 방식]")
    for result in sync_result.results[:5]:
        print(f"   {result}")

    print("\n[비동기 방식]")
    for result in async_result.results[:5]:
        print(f"   {result}")

    # 5. 결론
    print_section("💡 결론")
    print("""
    비동기 프로그래밍이 효과적인 경우:
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    ✓ 다수의 HTTP API 호출
    ✓ 데이터베이스 쿼리 (비동기 드라이버 사용 시)
    ✓ 파일 I/O 작업
    ✓ 웹소켓 연결 관리
    ✓ 외부 서비스와의 통신

    주의사항:
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    ⚠️  CPU 바운드 작업에는 multiprocessing 권장
    ⚠️  동기 코드와 혼합 시 주의 필요 (블로킹)
    ⚠️  디버깅이 상대적으로 복잡함
    ⚠️  모든 라이브러리가 비동기를 지원하지는 않음
    """)


def run() -> None:
    """
    프로그램 진입점

    asyncio.run()의 역할:
        1. 새로운 이벤트 루프 생성
        2. 전달된 코루틴 실행
        3. 완료 후 이벤트 루프 정리
    """
    if not HAS_HTTPX:
        print("❌ httpx 라이브러리를 설치해주세요: pip install httpx")
        return

    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n\n⚠️  사용자에 의해 중단되었습니다.")


if __name__ == "__main__":
    run()
