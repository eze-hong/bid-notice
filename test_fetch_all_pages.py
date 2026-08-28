#import time
#import math
#import xml.etree.ElementTree as ET
from unittest.mock import patch, MagicMock

import bid_notice_20260825 as bn # test srcipt

# ---------- 헬퍼 ----------

# 응답 객체 만들기
def make_resp(status=200, content=b"<response>ok</reponse>"):
    r = MagicMock()
    r.status_code = status
    r.headers = {"Content-Type": "application/xml"}
    r.content = content
    r.raise_for_status = lambda: None
    return r
    
def xml_page(total_count: int, items: list[str]) -> bytes:
    """items: ['A', 'B'] 형태로 넘기면 item/name 태그로 감싸서 XML bytes로 만듦"""
    items_xml = "".join(f"<item><name>{n}</name></item>" for n in items)
    return f"""
        <response>
            <body>
                <totalCount>{total_count}</totalCount>
                <items>{items_xml}</items>
            </body>
        </response>
    """.encode()
    
# JSON → ParseError 유발
BROKEN = b'{"resultCode":"99","resultMsg":"SERVICE_KEY_IS_NOT_REGISTERED_ERROR"}'

# page 당 1건 고정해 total_pages 계산 단순화
BASE_PARAMS = {"numOfRows": 1}


# ---------- 1. 전부 성공 ----------

def test_all_success():
    resp1 = make_resp(content=xml_page(total_count=2, items=["A"]))
    resp2 = make_resp(content=xml_page(total_count=2, items=["B"]))
    mock_get = MagicMock(side_effect=[resp1, resp2])
    
    with patch("time.sleep", return_value=None):
        items, failed_pages = bn.fetch_all_pages(
            session=MagicMock(get=mock_get), url="dummy", base_params=BASE_PARAMS
        )
    
    assert failed_pages == []
    assert len(items) == 2
    assert mock_get.call_count == 2


# ---------- 2. 1페이지 자체가 완전 실패 → RuntimeError ----------

def test_page1_fails_raises_runtime_error():
    broken = make_resp(content=BROKEN)
    mock_get = MagicMock(return_value=broken)
    
    with patch("time.sleep", return_value=None):
        try:
            bn.fetch_all_pages(session=MagicMock(get=mock_get), url="dummy", base_params=BASE_PARAMS)
            assert False, "RuntimeError가 발생했어야 함"
        except RuntimeError:
            pass
    
    assert mock_get.call_count == 3 # max_retry=3 소진 확인


# ---------- 3. 재시도 중간에 성공 (1,2차 실패 → 3차 성공) ----------

def test_retry_then_success():
    resp1_page1 = make_resp(content=xml_page(total_count=2, items=["A"]))
    broken = make_resp(content=BROKEN)
    resp_page2_ok = make_resp(content=xml_page(total_count=2, items=["B"]))
    
    # 호출 순서: 1페이지(성공) → 2페이지 1차(실패) → 2페이지 2차(실패) → 2페이지 3차(성공)
    mock_get = MagicMock(side_effect=[resp1_page1, broken, broken, resp_page2_ok])
    
    with patch("time.sleep", return_value=None):
        items, failed_pages = bn.fetch_all_pages(
            session=MagicMock(get=mock_get), url="dummy", base_params=BASE_PARAMS
        )
    
    assert failed_pages == []
    assert len(items) == 2
    assert mock_get.call_count == 4
    

# ---------- 4. 재시도 다 소진 → 해당 페이지만 스킵, 다음 페이지는 계속 진행 ----------

def test_retry_exhausted_page_skipped_continues_next():
    resp_page1 = make_resp(content=xml_page(total_count=3, items=["A"]))
    broken = make_resp(content=BROKEN)
    resp_page3 = make_resp(content=xml_page(total_count=3, items=["C"]))
    
    # 1page 성공 → 2page 3번 실패 → 3page 성공
    mock_get = MagicMock(side_effect=[resp_page1, broken, broken, broken, resp_page3])
    
    with patch("time.sleep", return_value=None):
        items, failed_pages = bn.fetch_all_pages(
            session=MagicMock(get=mock_get), url="dummy", base_params=BASE_PARAMS
        )
        
    assert failed_pages == [2]      # 2page만 실패
    assert len(items) == 2
    assert mock_get.call_count == 5
    

# ---------- 5. 여러 페이지 실패 (순서 포함) ----------

def test_multiple_pages_fail_order_preserved():
    resp_page1 = make_resp(content=xml_page(total_count=5, items=["A"]))
    broken = make_resp(content=BROKEN)
    resp_page3 = make_resp(content=xml_page(total_count=5, items=["C"]))
    resp_page4 = make_resp(content=xml_page(total_count=5, items=["D"]))
    
    # 1p 성공 → 2p 실패 → 3p 성공 → 4p 성공  → 5p 실패
    mock_get = MagicMock(side_effect=[
        resp_page1, 
        broken, broken, broken, 
        resp_page3, resp_page4, 
        broken, broken, broken
    ])
    
    with patch("time.sleep", return_value=None):
        items, failed_pages = bn.fetch_all_pages(
            session=MagicMock(get=mock_get), url="dummy", base_params=BASE_PARAMS
        )
        
    assert failed_pages == [2, 5]
    assert len(items) == 3
    

# ---------- 6. RequestException (네트워크 에러)도 동일하게 재시도 타는지 ----------

def test_request_exception_also_retried():
    import requests
    resp_ok = make_resp(content=xml_page(total_count=1, items=["A"]))
    
    mock_get = MagicMock(side_effect=[
        requests.exceptions.ConnectionError("network down"),
        requests.exceptions.ConnectionError("network down"),
        resp_ok,
    ])
    
    with patch("time.sleep", return_value=None):
        items, failed_pages = bn.fetch_all_pages(
            session=MagicMock(get=mock_get), url="dummy", base_params=BASE_PARAMS
        )
        
    assert failed_pages == []
    assert len(items) == 1
    assert mock_get.call_count == 3


# ---------- 7. 재시도 사이마다 sleep이 호출되는지 ----------

def test_sleep_called_between_retries():
    broken = make_resp(content=BROKEN)
    mock_get = MagicMock(return_value=broken)
    
    with patch("time.sleep", return_value=None) as mock_sleep:
        try:
            bn.fetch_all_pages(session=MagicMock(get=mock_get), url="dummy", base_params=BASE_PARAMS)
        except RuntimeError:
            pass
    
    # max_retry=3이면 sleep은 2번 호출됨
    assert mock_sleep.call_count == 2
    

def test_notify_teams():
    bn.notify_teams(
        "🧪 Teams Webhook 테스트",
        "테스트 메시지입니다."
    )