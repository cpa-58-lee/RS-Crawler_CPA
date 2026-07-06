"""
회생사건 자동조회 GUI (app.py)
- crawler_v10.py 엔진을 그대로 사용
- 엑셀 표 붙여넣기 입력 (경로 설정 불필요)
- 실시간 진행률 / 성공·오류 카운트 / 터미널 로그
설치: pip install PySide6
      pip install selenium openpyxl 2captcha-python webdriver-manager python-dotenv pandas
실행: python app.py
"""

import sys
import os
import time
import traceback
from datetime import datetime

from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtGui import QFont, QColor, QPalette
from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QLabel, QLineEdit, QPushButton,
    QPlainTextEdit, QTableWidget, QTableWidgetItem, QProgressBar,
    QHBoxLayout, QVBoxLayout, QGridLayout, QFrame, QHeaderView, QFileDialog,
    QMessageBox, QSizePolicy,
)

import crawler_v10 as engine


# ============================================================
# 색상 팔레트 (다크 테마 데모용)
# ============================================================
BG        = "#0f1117"
PANEL     = "#171a21"
PANEL2    = "#1d2129"
BORDER    = "#2a2f3a"
TEXT      = "#e6e8ec"
TEXT_SUB  = "#8b93a1"
ACCENT    = "#3b82f6"
GREEN     = "#22c55e"
AMBER     = "#f59e0b"
RED       = "#ef4444"
TERM_BG   = "#0a0d13"
TERM_TX   = "#c9d1d9"


# ============================================================
# 크롤링 워커 (백그라운드 스레드)
# ============================================================
class CrawlWorker(QObject):
    log        = Signal(str)
    progress   = Signal(int, int)   # done, total
    case_done  = Signal(int, int, dict, float)  # idx, total, result, elapsed
    finished   = Signal(dict)       # summary
    failed     = Signal(str)

    def __init__(self, cases, api_key, btpr_nm, output_path):
        super().__init__()
        self.cases = cases
        self.api_key = api_key
        self.btpr_nm = btpr_nm
        self.output_path = output_path
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        try:
            _, summary = engine.run_crawl(
                self.cases, self.api_key, self.btpr_nm, self.output_path,
                on_log=lambda m: self.log.emit(m),
                on_progress=lambda d, t: self.progress.emit(d, t),
                on_case_done=lambda i, t, r, e: self.case_done.emit(i, t, r, e),
                should_stop=lambda: self._stop,
            )
            self.finished.emit(summary)
        except Exception:
            self.failed.emit(traceback.format_exc())


# ============================================================
# 메인 윈도우
# ============================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("회생사건 자동조회 · v10")
        self.resize(1080, 680)
        self.thread = None
        self.worker = None
        self.n_ok = self.n_none = self.n_err = 0

        self._build_ui()
        self._apply_style()

    # ---------------- UI 구성 ----------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── 헤더 바 ──
        header = QFrame()
        header.setObjectName("header")
        hl = QHBoxLayout(header)
        hl.setContentsMargins(18, 12, 18, 12)

        logo = QLabel("⚖")
        logo.setObjectName("logo")
        title_box = QVBoxLayout()
        title_box.setSpacing(0)
        t1 = QLabel("회생사건 자동조회")
        t1.setObjectName("title")
        t2 = QLabel("대법원 나의사건검색 · v10")
        t2.setObjectName("subtitle")
        title_box.addWidget(t1)
        title_box.addWidget(t2)

        hl.addWidget(logo)
        hl.addLayout(title_box)
        hl.addStretch()

        key_lbl = QLabel("인증 키")
        key_lbl.setObjectName("keylabel")
        self.key_input = QLineEdit()
        self.key_input.setEchoMode(QLineEdit.Password)
        self.key_input.setPlaceholderText("API 키 입력")
        self.key_input.setFixedWidth(190)
        # 키 로드 우선순위: ① 빌드 시 주입된 내장 키 → ② 환경변수 → ③ .env 파일
        env_key = ""
        try:
            import _key  # 빌드 시 GitHub Secrets로부터 생성됨 (저장소엔 없음)
            env_key = getattr(_key, "VERIFY_API_KEY", "")
        except Exception:
            pass
        if not env_key:
            env_key = os.getenv("VERIFY_API_KEY", "") or os.getenv("TWOCAPTCHA_API_KEY", "")
        if not env_key:
            try:
                from dotenv import dotenv_values
                _v = {**dotenv_values(".env_verify"), **dotenv_values(".env_capcha")}
                env_key = _v.get("VERIFY_API_KEY", "") or _v.get("TWOCAPTCHA_API_KEY", "")
            except Exception:
                pass
        if env_key:
            self.key_input.setText(env_key)
        hl.addWidget(key_lbl)
        hl.addWidget(self.key_input)

        root.addWidget(header)

        # ── 본문 (좌/우) ──
        body = QHBoxLayout()
        body.setContentsMargins(18, 16, 18, 16)
        body.setSpacing(16)

        body.addWidget(self._build_left_panel(), 5)
        body.addWidget(self._build_right_panel(), 5)
        root.addLayout(body)

        # ── 상태바 ──
        self.statusBar().showMessage("표를 붙여넣고 조회를 시작하세요")

    def _build_left_panel(self):
        panel = QFrame()
        panel.setObjectName("panel")
        v = QVBoxLayout(panel)
        v.setContentsMargins(16, 14, 16, 14)
        v.setSpacing(10)

        cap = QLabel("📋  엑셀 표 붙여넣기")
        cap.setObjectName("caption")
        v.addWidget(cap)

        hint = QLabel("엑셀에서 [일련번호·차주명·법원·사건번호] 열을 복사해 아래에 붙여넣으세요 (Ctrl+V)")
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        v.addWidget(hint)

        self.paste_box = QPlainTextEdit()
        self.paste_box.setObjectName("pastebox")
        self.paste_box.setPlaceholderText(
            "EY-001\t홍길동\t대구회생법원\t2026회단102\n"
            "EY-002\t김철수\t대구회생법원\t2026간회단1012\n..."
        )
        self.paste_box.textChanged.connect(self._on_paste_changed)
        v.addWidget(self.paste_box, 1)

        # 미리보기 테이블
        self.preview = QTableWidget(0, 4)
        self.preview.setObjectName("preview")
        self.preview.setHorizontalHeaderLabels(["일련번호", "차주명", "법원", "사건번호"])
        self.preview.verticalHeader().setVisible(False)
        self.preview.setEditTriggers(QTableWidget.NoEditTriggers)
        self.preview.setSelectionMode(QTableWidget.NoSelection)
        hh = self.preview.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        hh.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.preview.setFixedHeight(150)
        v.addWidget(self.preview)

        # 당사자명 + 카운트
        row = QHBoxLayout()
        row.addWidget(QLabel("당사자명"))
        self.btpr_input = QLineEdit("아이")
        self.btpr_input.setFixedWidth(120)
        row.addWidget(self.btpr_input)
        self.count_lbl = QLabel("· 0건 인식됨")
        self.count_lbl.setObjectName("hint")
        row.addWidget(self.count_lbl)
        row.addStretch()
        v.addLayout(row)

        # 버튼
        btn_row = QHBoxLayout()
        self.start_btn = QPushButton("▶  조회 시작")
        self.start_btn.setObjectName("primary")
        self.start_btn.clicked.connect(self._on_start)
        self.stop_btn = QPushButton("■ 중단")
        self.stop_btn.setObjectName("ghost")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._on_stop)
        btn_row.addWidget(self.start_btn, 1)
        btn_row.addWidget(self.stop_btn)
        v.addLayout(btn_row)

        return panel

    def _build_right_panel(self):
        panel = QFrame()
        panel.setObjectName("panel")
        v = QVBoxLayout(panel)
        v.setContentsMargins(16, 14, 16, 14)
        v.setSpacing(10)

        top = QHBoxLayout()
        cap = QLabel("📡  진행 상황")
        cap.setObjectName("caption")
        top.addWidget(cap)
        top.addStretch()
        self.prog_lbl = QLabel("0 / 0")
        self.prog_lbl.setObjectName("proglabel")
        top.addWidget(self.prog_lbl)
        v.addLayout(top)

        self.progress = QProgressBar()
        self.progress.setObjectName("progress")
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(8)
        v.addWidget(self.progress)

        self.now_lbl = QLabel("대기 중…")
        self.now_lbl.setObjectName("hint")
        v.addWidget(self.now_lbl)

        # 카운트 카드
        cards = QHBoxLayout()
        cards.setSpacing(8)
        self.card_ok   = self._make_card("성공", GREEN)
        self.card_none = self._make_card("사건없음", AMBER)
        self.card_err  = self._make_card("오류", RED)
        cards.addWidget(self.card_ok["frame"])
        cards.addWidget(self.card_none["frame"])
        cards.addWidget(self.card_err["frame"])
        v.addLayout(cards)

        # 터미널 로그
        self.term = QPlainTextEdit()
        self.term.setObjectName("terminal")
        self.term.setReadOnly(True)
        v.addWidget(self.term, 1)

        return panel

    def _make_card(self, label, color):
        f = QFrame()
        f.setObjectName("card")
        f.setStyleSheet(
            f"#card {{ background:{PANEL2}; border-radius:8px; }}"
        )
        lay = QVBoxLayout(f)
        lay.setContentsMargins(12, 8, 12, 8)
        lay.setSpacing(0)
        num = QLabel("0")
        num.setStyleSheet(f"color:{color}; font-size:22px; font-weight:600;")
        txt = QLabel(label)
        txt.setStyleSheet(f"color:{color}; font-size:11px;")
        lay.addWidget(num)
        lay.addWidget(txt)
        return {"frame": f, "num": num}

    # ---------------- 로직 ----------------
    def _on_paste_changed(self):
        text = self.paste_box.toPlainText()
        try:
            cases = engine.parse_pasted_table(text)
        except Exception:
            cases = []
        self._cases_cache = cases
        self.count_lbl.setText(f"· {len(cases)}건 인식됨")

        self.preview.setRowCount(0)
        for c in cases[:200]:
            r = self.preview.rowCount()
            self.preview.insertRow(r)
            for col, key in enumerate(["일련번호", "차주명", "법원", "사건번호"]):
                self.preview.setItem(r, col, QTableWidgetItem(c[key]))

    def _on_start(self):
        cases = getattr(self, "_cases_cache", [])
        api_key = self.key_input.text().strip()
        btpr = self.btpr_input.text().strip()

        if not cases:
            QMessageBox.warning(self, "입력 필요", "먼저 엑셀 표를 붙여넣어 주세요.")
            return
        if not api_key:
            QMessageBox.warning(self, "인증 키 필요",
                                "인증 API 키를 입력해 주세요.\n(우측 상단 입력칸)")
            return
        if not btpr:
            QMessageBox.warning(self, "당사자명 필요", "당사자명을 입력해 주세요.")
            return

        # 저장 경로
        default_name = f"회생사건_조회결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "결과 저장 위치", default_name,
                                              "Excel (*.xlsx)")
        if not path:
            return
        self.output_path = path

        # 초기화
        self.n_ok = self.n_none = self.n_err = 0
        self._update_cards()
        self.term.clear()
        self.progress.setValue(0)
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.paste_box.setEnabled(False)

        # 워커 스레드 시작
        self.thread = QThread()
        self.worker = CrawlWorker(cases, api_key, btpr, path)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.log.connect(self._on_log)
        self.worker.progress.connect(self._on_progress)
        self.worker.case_done.connect(self._on_case_done)
        self.worker.finished.connect(self._on_finished)
        self.worker.failed.connect(self._on_failed)
        self.thread.start()

    def _on_stop(self):
        if self.worker:
            self.worker.stop()
            self.now_lbl.setText("중단 요청됨… 현재 사건 완료 후 정지")
            self.stop_btn.setEnabled(False)

    def _on_log(self, msg):
        self.term.appendPlainText(msg)
        self.term.verticalScrollBar().setValue(
            self.term.verticalScrollBar().maximum())

    def _on_progress(self, done, total):
        pct = int(done / total * 100) if total else 0
        self.progress.setMaximum(100)
        self.progress.setValue(pct)
        self.prog_lbl.setText(f"{done} / {total}")

    def _on_case_done(self, idx, total, result, elapsed):
        res = result.get("결과", "")
        if res == "조회성공":
            self.n_ok += 1
        elif res == "사건 없음":
            self.n_none += 1
        else:
            self.n_err += 1
        self._update_cards()
        court = result.get("관할법원", "")
        case = result.get("사건번호", "")
        self.now_lbl.setText(f"{court} {case} 완료 ({elapsed:.1f}초)")

    def _update_cards(self):
        self.card_ok["num"].setText(str(self.n_ok))
        self.card_none["num"].setText(str(self.n_none))
        self.card_err["num"].setText(str(self.n_err))

    def _on_finished(self, summary):
        self._cleanup_thread()
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.paste_box.setEnabled(True)
        self.now_lbl.setText("✅ 완료")
        path = summary.get("path")
        self.statusBar().showMessage(
            f"완료 — 성공 {summary['ok']} / 사건없음 {summary['none']} / 오류 {summary['err']}"
        )
        msg = (f"조회 완료\n\n성공: {summary['ok']}건\n사건없음: {summary['none']}건\n"
               f"오류: {summary['err']}건\n\n저장: {path}")
        box = QMessageBox(self)
        box.setWindowTitle("완료")
        box.setText(msg)
        open_btn = box.addButton("폴더 열기", QMessageBox.AcceptRole)
        box.addButton("닫기", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() == open_btn and path:
            self._open_folder(os.path.dirname(os.path.abspath(path)))

    def _on_failed(self, tb):
        self._cleanup_thread()
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.paste_box.setEnabled(True)
        self.now_lbl.setText("❌ 오류 발생")
        self._on_log("\n[치명적 오류]\n" + tb)
        QMessageBox.critical(self, "오류", "크롤링 중 오류가 발생했습니다.\n로그를 확인하세요.")

    def _cleanup_thread(self):
        if self.thread:
            self.thread.quit()
            self.thread.wait()
            self.thread = None
            self.worker = None

    def _open_folder(self, folder):
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception:
            pass

    def closeEvent(self, event):
        if self.worker:
            self.worker.stop()
        self._cleanup_thread()
        event.accept()

    # ---------------- 스타일 ----------------
    def _apply_style(self):
        self.setStyleSheet(f"""
            QMainWindow {{ background:{BG}; }}
            QWidget {{ color:{TEXT}; font-family:'Malgun Gothic','맑은 고딕',sans-serif; font-size:13px; }}
            #header {{ background:{PANEL}; border-bottom:1px solid {BORDER}; }}
            #logo {{ background:{ACCENT}; color:white; font-size:18px; border-radius:8px;
                     min-width:32px; max-width:32px; min-height:32px; max-height:32px;
                     qproperty-alignment: AlignCenter; }}
            #title {{ font-size:15px; font-weight:600; color:{TEXT}; }}
            #subtitle {{ font-size:11px; color:{TEXT_SUB}; }}
            #keylabel {{ color:{TEXT_SUB}; font-size:12px; }}
            #panel {{ background:{PANEL}; border:1px solid {BORDER}; border-radius:12px; }}
            #caption {{ font-size:13px; font-weight:600; color:{TEXT}; }}
            #hint {{ color:{TEXT_SUB}; font-size:11px; }}
            #proglabel {{ color:{ACCENT}; font-size:12px; font-weight:600; }}
            QLineEdit {{ background:{PANEL2}; border:1px solid {BORDER}; border-radius:8px;
                         padding:6px 10px; color:{TEXT}; }}
            QLineEdit:focus {{ border:1px solid {ACCENT}; }}
            #pastebox {{ background:{PANEL2}; border:1px solid {BORDER}; border-radius:8px;
                         padding:8px; color:{TEXT}; font-family:'Consolas','D2Coding',monospace;
                         font-size:12px; }}
            #pastebox:focus {{ border:1px solid {ACCENT}; }}
            #preview {{ background:{PANEL2}; border:1px solid {BORDER}; border-radius:8px;
                        gridline-color:{BORDER}; font-size:11px; }}
            #preview QHeaderView::section {{ background:{PANEL}; color:{TEXT_SUB};
                        border:none; border-bottom:1px solid {BORDER}; padding:5px; }}
            #terminal {{ background:{TERM_BG}; border:1px solid {BORDER}; border-radius:8px;
                         color:{TERM_TX}; font-family:'Consolas','D2Coding',monospace;
                         font-size:11px; padding:8px; }}
            QProgressBar {{ background:{PANEL2}; border:none; border-radius:4px; }}
            QProgressBar::chunk {{ background:{ACCENT}; border-radius:4px; }}
            QPushButton {{ padding:8px 14px; border-radius:8px; font-size:13px; }}
            #primary {{ background:{ACCENT}; color:white; font-weight:600; border:none; }}
            #primary:hover {{ background:#2f6fd6; }}
            #primary:disabled {{ background:{PANEL2}; color:{TEXT_SUB}; }}
            #ghost {{ background:transparent; color:{TEXT_SUB}; border:1px solid {BORDER}; }}
            #ghost:hover {{ background:{PANEL2}; }}
            #ghost:disabled {{ color:{BORDER}; }}
            QStatusBar {{ background:{PANEL}; color:{TEXT_SUB}; }}
            QScrollBar:vertical {{ background:transparent; width:8px; }}
            QScrollBar::handle:vertical {{ background:{BORDER}; border-radius:4px; }}
        """)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
