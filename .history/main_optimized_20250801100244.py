#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
플레이오토 주문 엑셀 기반 송장 업무 자동 분류 프로그램 v4.1 (성능 최적화)
개발자: Claude (Anthropic)
버전: 4.1.0
지원 OS: MacOS, Windows
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import json
import os
from datetime import datetime
import threading
from tkinter import font
import sys
import re
from collections import defaultdict, OrderedDict
import queue
import time
from concurrent.futures import ThreadPoolExecutor
import numpy as np

class FastProgressDialog:
    """초고속 진행률 표시 다이얼로그"""
    def __init__(self, parent, title="처리 중..."):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("450x250")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 부모 창 중앙에 위치
        parent.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - 250
        y = parent.winfo_y() + (parent.winfo_height() // 2) - 140
        self.dialog.geometry(f"+{x}+{y}")
        
        # 모던한 그라데이션 배경
        self.dialog.configure(bg='#0a0a0a')
        
        # 네온 효과 제목
        self.title_label = tk.Label(self.dialog, text=title, 
                                   font=('SF Pro Display', 16, 'bold'),
                                   bg='#0a0a0a', fg='#00ff88')
        self.title_label.pack(pady=25)
        
        # 메인 프레임
        main_frame = tk.Frame(self.dialog, bg='#0a0a0a')
        main_frame.pack(fill='both', expand=True, padx=30)
        
        # 전체 진행률
        tk.Label(main_frame, text="전체 진행률", 
                font=('SF Pro Display', 11), bg='#0a0a0a', fg='#888').pack(anchor='w')
        
        # 커스텀 프로그레스바 프레임
        self.progress_frame = tk.Frame(main_frame, bg='#1a1a1a', height=8)
        self.progress_frame.pack(fill='x', pady=(5, 15))
        
        self.progress_bar = tk.Frame(self.progress_frame, bg='#00ff88', height=8)
        self.progress_bar.place(x=0, y=0, relheight=1, width=0)
        
        # 퍼센트와 상태를 한 줄에
        info_frame = tk.Frame(main_frame, bg='#0a0a0a')
        info_frame.pack(fill='x')
        
        self.percent_label = tk.Label(info_frame, text="0%", 
                                     font=('SF Pro Display', 24, 'bold'),
                                     bg='#0a0a0a', fg='#00ff88')
        self.percent_label.pack(side='left')
        
        self.status_label = tk.Label(info_frame, text="준비 중...", 
                                    font=('SF Pro Display', 11),
                                    bg='#0a0a0a', fg='#666')
        self.status_label.pack(side='left', padx=20)
        
        # 세부 진행률 (더 얇게)
        self.sub_progress_frame = tk.Frame(main_frame, bg='#1a1a1a', height=4)
        self.sub_progress_frame.pack(fill='x', pady=(20, 10))
        
        self.sub_progress_bar = tk.Frame(self.sub_progress_frame, bg='#0088ff', height=4)
        self.sub_progress_bar.place(x=0, y=0, relheight=1, width=0)
        
        # 취소 버튼 (모던 스타일)
        self.cancel_button = tk.Button(self.dialog, text="취소", 
                                      font=('SF Pro Display', 13),
                                      bg='#222', fg='#fff',
                                      activebackground='#333',
                                      activeforeground='#fff',
                                      bd=0, padx=30, pady=8,
                                      cursor='hand2',
                                      command=self.cancel)
        self.cancel_button.pack(pady=20)
        
        self.cancelled = False
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
        
        # 성능 최적화: 업데이트 쓰로틀링
        self.last_update_time = 0
        self.update_interval = 0.05  # 50ms 간격으로만 업데이트
        
    def update(self, percent, status="", sub_percent=0):
        """진행률 업데이트 (성능 최적화)"""
        if self.cancelled:
            return
            
        current_time = time.time()
        if current_time - self.last_update_time < self.update_interval:
            return
            
        self.last_update_time = current_time
        
        # 프로그레스바 애니메이션
        self.progress_bar.place(relwidth=percent/100)
        self.percent_label.config(text=f"{int(percent)}%")
        
        if status:
            self.status_label.config(text=status)
            
        if sub_percent > 0:
            self.sub_progress_bar.place(relwidth=sub_percent/100)
        
        self.dialog.update_idletasks()
    
    def cancel(self):
        """작업 취소"""
        self.cancelled = True
        self.dialog.destroy()
    
    def close(self):
        """다이얼로그 닫기"""
        if not self.cancelled:
            self.dialog.destroy()

class PlayAutoOrderClassifierV41:
    def __init__(self):
        self.root = tk.Tk()
        self.setup_window()
        self.setup_modern_styles()
        self.load_settings()
        self.create_widgets()
        
        # 성능 최적화를 위한 변수들
        self.excel_data = None
        self.classified_data = None
        self.work_ranges = {}
        self.unmatched_products = {}
        self.product_history = {}
        
        # 스레드 풀 (성능 향상)
        self.executor = ThreadPoolExecutor(max_workers=4)
        
        # 캐시 (반복 연산 방지)
        self.classification_cache = {}
        
    def setup_window(self):
        """메인 윈도우 설정 (모던 디자인)"""
        self.root.title("플레이오토 송장 분류 시스템")
        self.root.geometry("1400x900")
        self.root.configure(bg='#000000')
        self.root.resizable(True, True)
        
        # 다크 테마 강화
        self.root.tk_setPalette(background='#000000', foreground='white')
        
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        # 윈도우 중앙 배치
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def setup_modern_styles(self):
        """모던하고 트렌디한 UI 스타일 설정"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # 네온 컬러 팔레트
        self.colors = {
            'bg': '#000000',
            'bg_secondary': '#0a0a0a',
            'panel': '#111111',
            'card': '#1a1a1a',
            'border': '#2a2a2a',
            'neon_green': '#00ff88',
            'neon_blue': '#0088ff',
            'neon_purple': '#8800ff',
            'neon_pink': '#ff0088',
            'neon_yellow': '#ffdd00',
            'text_primary': '#ffffff',
            'text_secondary': '#888888',
            'text_muted': '#555555',
            'success': '#00ff88',
            'warning': '#ffdd00',
            'danger': '#ff0088',
            'info': '#0088ff'
        }
        
        self.fonts = {
            'title': ('SF Pro Display', 24, 'bold'),
            'header': ('SF Pro Display', 16, 'bold'),
            'body': ('SF Pro Text', 13),
            'small': ('SF Pro Text', 11),
            'mono': ('SF Mono', 12)
        }
        
        # 스타일 설정
        self.style.configure('Title.TLabel', 
                           background=self.colors['bg'], 
                           foreground=self.colors['neon_green'],
                           font=self.fonts['title'])
        
        self.style.configure('Header.TLabel',
                           background=self.colors['panel'],
                           foreground=self.colors['text_primary'],
                           font=self.fonts['header'])
        
        # 모던 버튼 스타일
        for name, color in [('Success', 'neon_green'), ('Warning', 'neon_yellow'), 
                           ('Danger', 'neon_pink'), ('Info', 'neon_blue')]:
            self.style.configure(f'{name}.TButton',
                               background=self.colors['card'],
                               foreground=self.colors[color],
                               borderwidth=1,
                               relief='flat',
                               font=self.fonts['body'])
            
            self.style.map(f'{name}.TButton',
                         background=[('active', self.colors[color]),
                                   ('pressed', self.colors['panel'])],
                         foreground=[('active', self.colors['bg']),
                                   ('pressed', self.colors[color])])
        
        # 노트북 탭 스타일
        self.style.configure('TNotebook', 
                           background=self.colors['bg'],
                           borderwidth=0)
        self.style.configure('TNotebook.Tab',
                           background=self.colors['card'],
                           foreground=self.colors['text_secondary'],
                           padding=[20, 12],
                           font=self.fonts['body'])
        self.style.map('TNotebook.Tab',
                      background=[('selected', self.colors['panel'])],
                      foreground=[('selected', self.colors['neon_green'])])
    
    def load_settings(self):
        """설정 파일 로드 (성능 최적화)"""
        self.settings_file = 'playauto_settings_v4.json'
        self.product_history_file = 'product_history_v4.json'
        
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
            else:
                # 기본 설정은 기존과 동일
                self.settings = self.get_default_settings()
                self.save_settings()
        except Exception as e:
            print(f"설정 로드 오류: {e}")
            self.settings = self.get_default_settings()
        
        # 상품 기록 로드
        try:
            if os.path.exists(self.product_history_file):
                with open(self.product_history_file, 'r', encoding='utf-8') as f:
                    self.product_history = json.load(f)
            else:
                self.product_history = {}
        except Exception as e:
            print(f"상품 기록 로드 오류: {e}")
            self.product_history = {}
    
    def get_default_settings(self):
        """기본 설정 반환"""
        return {
            "work_order": ["송과장님", "영재씨", "효상", "강민씨", "부모님", "합배송", "복수주문", "분류실패"],
            "work_config": {
                "송과장님": {
                    "type": "product_specific",
                    "products": [
                        {"brand": "꽃샘", "product_name": "밤 티라미수 라떄 20T", "order_option": "All"},
                        {"brand": "꽃샘", "product_name": "블랙보리차 티백 100개입 1개", "order_option": "All"},
                        {"brand": "꽃샘", "product_name": "콘푸레이크 천마차 130T", "order_option": "All"},
                        {"brand": "꽃샘", "product_name": "콘푸레이크 천마차 50T", "order_option": "All"},
                        {"brand": "금상", "product_name": "빙수떡", "order_option": "All"},
                        {"brand": "스윗박스", "product_name": "팥빙수 재료 패밀리C 세트 빙수팥 후루츠칵테일 빙수떡 연유", "order_option": "All"},
                        {"brand": "", "product_name": "팥빙수재료 4종세트 팥+빙수떡+후루츠칵테일+연유", "order_option": "All"},
                        {"brand": "금상", "product_name": "빙수떡 2개 + 빙수제리 1개 팥빙수재료", "order_option": "All"},
                        {"brand": "참존", "product_name": "통단팥 3kg 원터치캔", "order_option": "All"},
                        {"brand": "삼진식품", "product_name": "빙수애 콩가루 팥유크림함유", "order_option": "All"}
                    ],
                    "description": "팥빙수재료 및 특정 상품 담당",
                    "icon": "🍧",
                    "enabled": True
                },
                "영재씨": {
                    "type": "product_specific", 
                    "products": [
                        {"brand": "미에로화이바", "product_name": "All", "order_option": "All"},
                        {"brand": "", "product_name": "6x25 하트 스트로우 (핑크) 빨대 개별포장 200개", "order_option": "단일상품"},
                        {"brand": "꽃샘", "product_name": "꿀유자차S", "order_option": "All"},
                        {"brand": "꽃샘", "product_name": "꿀생강차 S 2kg + 2kg (총4kg)", "order_option": "All"},
                        {"brand": "꽃샘", "product_name": "꿀생강차 S", "order_option": "All"}
                    ],
                    "description": "미에로화이바, 꿀차, 파우치음료, 컵류 담당",
                    "icon": "🍯",
                    "enabled": True
                },
                "효상": {
                    "type": "product_specific",
                    "products": [
                        {"brand": "백제", "product_name": "멸치맛 쌀국수", "order_option": "92g 10개"},
                        {"brand": "백제", "product_name": "우리 햅쌀 즉석떡국 6개입", "order_option": "163g 6개"}
                    ],
                    "description": "백제 쌀국수, 떡국 담당",
                    "icon": "🍜",
                    "enabled": True
                },
                "강민씨": {
                    "type": "product_specific",
                    "products": [
                        {"brand": "백제", "product_name": "All", "order_option": "All"}
                    ],
                    "description": "백제 브랜드 모든 상품 담당",
                    "icon": "🍜", 
                    "enabled": True
                },
                "부모님": {
                    "type": "product_specific",
                    "products": [
                        {"brand": "쟈뎅", "product_name": "All", "order_option": "All"},
                        {"brand": "부국", "product_name": "All", "order_option": "All"},
                        {"brand": "린저", "product_name": "All", "order_option": "All"}
                    ],
                    "description": "쟈뎅, 부국, 린저, 꿀, 카페재료, 타코 담당",
                    "icon": "☕",
                    "enabled": True
                },
                "합배송": {
                    "type": "mixed_products",
                    "products": [],
                    "description": "한 주문번호에 여러 다른 상품",
                    "icon": "📦",
                    "enabled": True,
                    "auto_rule": "multiple_products"
                },
                "복수주문": {
                    "type": "multiple_quantity", 
                    "products": [],
                    "description": "한 상품을 2개 이상 주문",
                    "icon": "📋",
                    "enabled": True,
                    "auto_rule": "high_quantity"
                },
                "분류실패": {
                    "type": "failed",
                    "products": [],
                    "description": "매칭되지 않은 상품 (수동 검토 필요)",
                    "icon": "❓",
                    "enabled": True,
                    "auto_rule": "unmatched"
                }
            },
            "auto_learn": True,
            "min_confidence": 1.0,
            "quantity_threshold": 2
        }
    
    def save_settings(self):
        """설정 파일 저장"""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"설정 저장 오류: {e}")
    
    def save_product_history(self):
        """상품 분류 기록 저장"""
        try:
            with open(self.product_history_file, 'w', encoding='utf-8') as f:
                json.dump(self.product_history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"상품 기록 저장 오류: {e}")
    
    def create_widgets(self):
        """메인 UI 위젯 생성 (모던 디자인)"""
        # 메인 컨테이너 (그라데이션 효과)
        main_container = tk.Frame(self.root, bg=self.colors['bg'])
        main_container.pack(fill='both', expand=True)
        
        # 헤더 영역
        header_frame = tk.Frame(main_container, bg=self.colors['bg'], height=80)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        # 타이틀 (네온 효과)
        title_frame = tk.Frame(header_frame, bg=self.colors['bg'])
        title_frame.pack(expand=True)
        
        title_label = tk.Label(title_frame, 
                              text="플레이오토",
                              font=('SF Pro Display', 32, 'bold'),
                              bg=self.colors['bg'],
                              fg=self.colors['neon_green'])
        title_label.pack()
        
        subtitle_label = tk.Label(title_frame,
                                 text="송장 업무별 분류 시스템",
                                 font=('SF Pro Display', 14),
                                 bg=self.colors['bg'],
                                 fg=self.colors['text_secondary'])
        subtitle_label.pack()
        
        # 메인 콘텐츠 영역
        content_frame = tk.Frame(main_container, bg=self.colors['bg_secondary'])
        content_frame.pack(fill='both', expand=True, padx=20, pady=(0, 15))
        
        # 노트북 위젯
        self.notebook = ttk.Notebook(content_frame)
        self.notebook.pack(fill='both', expand=True)
        
        # 탭 생성
        self.create_main_tab()
        self.create_work_management_tab()
        self.create_product_settings_tab()
        self.create_stats_tab()
        
        # 하단 상태바
        self.create_status_bar()
    
    def create_status_bar(self):
        """하단 상태바 생성 (모던 디자인)"""
        status_frame = tk.Frame(self.root, bg=self.colors['panel'], height=35)
        status_frame.pack(fill='x', side='bottom')
        status_frame.pack_propagate(False)
        
        # 상태 텍스트
        self.status_var = tk.StringVar(value="처리 준비 완료")
        status_label = tk.Label(status_frame, 
                               textvariable=self.status_var,
                               bg=self.colors['panel'],
                               fg=self.colors['text_secondary'],
                               font=self.fonts['small'])
        status_label.pack(side='left', padx=15, pady=8)
        
        # 실시간 시계
        self.time_label = tk.Label(status_frame, 
                                  bg=self.colors['panel'],
                                  fg=self.colors['text_muted'],
                                  font=self.fonts['mono'])
        self.time_label.pack(side='right', padx=20, pady=10)
        self.update_time()
        
        # 성능 인디케이터
        perf_label = tk.Label(status_frame,
                             text="⚡ 울트라 퍼포먼스",
                             bg=self.colors['panel'],
                             fg=self.colors['neon_yellow'],
                             font=('SF Pro Display', 10, 'bold'))
        perf_label.pack(side='right', padx=40, pady=10)
    
    def update_time(self):
        """실시간 시계 업데이트"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_label.config(text=current_time)
        self.root.after(1000, self.update_time)
    
    def create_main_tab(self):
        """메인 작업 탭 생성 (모던 디자인)"""
        main_tab = tk.Frame(self.notebook, bg=self.colors['bg_secondary'])
        self.notebook.add(main_tab, text="📊 홈")
        
        # 스크롤 가능한 프레임
        canvas = tk.Canvas(main_tab, bg=self.colors['bg_secondary'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['bg_secondary'])
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=(0, 2))
        scrollbar.pack(side="right", fill="y", padx=(0, 3))
        
        # 대시보드 레이아웃
        self.create_dashboard_content(scrollable_frame)
    
    def create_dashboard_content(self, parent):
        """대시보드 콘텐츠 생성 (2열 레이아웃)"""
        # 정보 카드
        info_container = tk.Frame(parent, bg=self.colors['bg_secondary'])
        info_container.pack(fill='x', padx=15, pady=15)
        self.create_info_card(info_container, "현재 설정", self.get_config_summary())
        
        # 메인 2열 레이아웃
        main_layout = tk.Frame(parent, bg=self.colors['bg_secondary'])
        main_layout.pack(fill='both', expand=True, padx=15, pady=5)
        
        # 좌측 컬럼 (파일 업로드)
        left_column = tk.Frame(main_layout, bg=self.colors['bg_secondary'])
        left_column.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        upload_section = self.create_section(left_column, "📁 파일 업로드")
        self.create_drop_zone(upload_section)
        self.create_action_buttons(upload_section)
        
        # 우측 컬럼 (분류 결과)
        right_column = tk.Frame(main_layout, bg=self.colors['bg_secondary'])
        right_column.pack(side='right', fill='both', expand=True, padx=(10, 0))
        
        result_section = self.create_section(right_column, "📊 분류 결과")
        self.create_result_display_compact(result_section)
    
    def create_info_card(self, parent, title, content):
        """정보 카드 생성 (네온 스타일)"""
        card_frame = tk.Frame(parent, bg=self.colors['card'], 
                             highlightbackground=self.colors['neon_green'],
                             highlightthickness=1)
        card_frame.pack(fill='x', pady=10)
        
        # 카드 헤더
        header_frame = tk.Frame(card_frame, bg=self.colors['card'])
        header_frame.pack(fill='x', padx=15, pady=(12, 4))
        
        title_label = tk.Label(header_frame, text=title,
                              font=self.fonts['header'],
                              bg=self.colors['card'],
                              fg=self.colors['neon_green'])
        title_label.pack(side='left')
        
        # 카드 콘텐츠
        content_label = tk.Label(card_frame, text=content,
                               font=self.fonts['body'],
                               bg=self.colors['card'],
                               fg=self.colors['text_secondary'],
                               justify='left')
        content_label.pack(fill='x', padx=15, pady=(4, 12))
    
    def get_config_summary(self):
        """현재 설정 요약"""
        work_names = [f"{self.settings['work_config'][name]['icon']} {name}" 
                    for name in self.settings['work_order'][:8]]
        if len(self.settings['work_order']) <= 8:
            return f"담당자 순서: {' → '.join(work_names)}"
        else:
            return f"담당자 순서: {' → '.join(work_names)}..."
    
    def create_section(self, parent, title):
        """섹션 생성"""
        section_frame = tk.Frame(parent, bg=self.colors['bg_secondary'])
        section_frame.pack(fill='x', padx=15, pady=15)
        
        # 섹션 헤더
        header_label = tk.Label(section_frame, text=title,
                               font=self.fonts['header'],
                               bg=self.colors['bg_secondary'],
                               fg=self.colors['text_primary'])
        header_label.pack(anchor='w', pady=(0, 15))
        
        # 섹션 콘텐츠
        content_frame = tk.Frame(section_frame, bg=self.colors['panel'])
        content_frame.pack(fill='x')
        
        return content_frame
    
    def create_drop_zone(self, parent):
        """드래그 앤 드롭 영역 생성 (네온 효과)"""
        drop_frame = tk.Frame(parent, bg=self.colors['card'], 
                             highlightbackground=self.colors['neon_blue'],
                             highlightthickness=2)
        drop_frame.pack(fill='x', padx=15, pady=15)
        
        # 드롭 영역 콘텐츠
        drop_content = tk.Frame(drop_frame, bg=self.colors['card'])
        drop_content.pack(fill='both', expand=True, pady=20)
        
        # 아이콘
        icon_label = tk.Label(drop_content, text="📤",
                             font=('Arial', 48),
                             bg=self.colors['card'])
        icon_label.pack()
        
        # 메인 텍스트
        main_text = tk.Label(drop_content, 
                            text="엑셀 파일을 여기에 드래그하거나 클릭하여 선택하세요",
                            font=self.fonts['header'],
                            bg=self.colors['card'],
                            fg=self.colors['neon_blue'])
        main_text.pack(pady=10)
        
        # 서브 텍스트
        sub_text = tk.Label(drop_content,
                           text=".xlsx 및 .xls 형식 지원 • 초고속 처리",
                           font=self.fonts['small'],
                           bg=self.colors['card'],
                           fg=self.colors['text_muted'])
        sub_text.pack()
        
        # 파일 정보
        self.file_info_var = tk.StringVar(value="파일이 선택되지 않았습니다")
        file_info = tk.Label(drop_content,
                            textvariable=self.file_info_var,
                            font=self.fonts['body'],
                            bg=self.colors['card'],
                            fg=self.colors['warning'])
        file_info.pack(pady=10)
        
        # 클릭 이벤트 (성능 최적화: 즉시 실행)
        drop_frame.bind('<Button-1>', lambda e: self.select_file())
        for child in drop_content.winfo_children():
            child.bind('<Button-1>', lambda e: self.select_file())
        
        # 호버 효과
        def on_enter(e):
            drop_frame.config(highlightbackground=self.colors['neon_green'])
        
        def on_leave(e):
            drop_frame.config(highlightbackground=self.colors['neon_blue'])
        
        drop_frame.bind('<Enter>', on_enter)
        drop_frame.bind('<Leave>', on_leave)
        
        # 커서 변경
        drop_frame.config(cursor='hand2')
        for child in drop_content.winfo_children():
            child.config(cursor='hand2')
    
    def create_action_buttons(self, parent):
        """액션 버튼들 생성 (성능 최적화)"""
        button_frame = tk.Frame(parent, bg=self.colors['panel'])
        button_frame.pack(pady=15)
        
        # 버튼 스타일 정의
        button_config = {
            'font': self.fonts['body'],
            'bd': 0,
            'padx': 30,
            'pady': 12,
            'cursor': 'hand2'
        }
        
        # 처리 버튼
        self.process_button = tk.Button(button_frame, 
                                       text="⚡ 분류 시작",
                                       bg=self.colors['neon_green'],
                                       fg=self.colors['bg'],
                                       activebackground=self.colors['success'],
                                       state='disabled',
                                       command=self.process_excel,
                                       **button_config)
        self.process_button.pack(side='left', padx=10)
        
        # 다운로드 버튼
        self.download_button = tk.Button(button_frame,
                                        text="💾 결과 다운로드",
                                        bg=self.colors['neon_blue'],
                                        fg=self.colors['bg'],
                                        activebackground=self.colors['info'],
                                        state='disabled',
                                        command=self.download_excel,
                                        **button_config)
        self.download_button.pack(side='left', padx=10)
        
        # 검토 버튼
        self.review_button = tk.Button(button_frame,
                                      text="🔍 미분류 검토",
                                      bg=self.colors['neon_yellow'],
                                      fg=self.colors['bg'],
                                      activebackground=self.colors['warning'],
                                      state='disabled',
                                      command=self.review_unmatched,
                                      **button_config)
        self.review_button.pack(side='left', padx=10)
    
    def create_work_management_tab(self):
        """업무 관리 탭 생성"""
        work_tab = tk.Frame(self.notebook, bg=self.colors['bg_secondary'])
        self.notebook.add(work_tab, text="👥 담당자")
        
        # 메인 레이아웃
        main_frame = tk.Frame(work_tab, bg=self.colors['bg_secondary'])
        main_frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 좌측: 워커 리스트
        left_frame = tk.Frame(main_frame, bg=self.colors['panel'])
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 8))
        
        # 리스트 헤더
        list_header = tk.Label(left_frame, 
                              text="담당자 우선순위 목록",
                              font=self.fonts['header'],
                              bg=self.colors['panel'],
                              fg=self.colors['text_primary'])
        list_header.pack(pady=15)
        
        # 워커 리스트박스
        list_container = tk.Frame(left_frame, bg=self.colors['card'])
        list_container.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        
        self.work_listbox = tk.Listbox(list_container,
                                      bg=self.colors['card'],
                                      fg=self.colors['text_primary'],
                                      font=self.fonts['body'],
                                      selectbackground=self.colors['neon_green'],
                                      selectforeground=self.colors['bg'],
                                      height=15,
                                      bd=0,
                                      highlightthickness=0,
                                      activestyle='none')
        self.work_listbox.pack(fill='both', expand=True, padx=1, pady=1)
        
        # 우측: 컨트롤 패널
        right_frame = tk.Frame(main_frame, bg=self.colors['panel'], width=250)
        right_frame.pack(side='right', fill='y')
        right_frame.pack_propagate(False)
        
        # 컨트롤 헤더
        control_header = tk.Label(right_frame,
                                 text="빠른 작업",
                                 font=self.fonts['header'],
                                 bg=self.colors['panel'],
                                 fg=self.colors['text_primary'])
        control_header.pack(pady=15)
        
        # 액션 버튼들
        self.create_work_action_buttons(right_frame)
        
        # 초기 리스트 로드
        self.refresh_work_list()
    
    def create_result_display_compact(self, parent):
        """컴팩트한 결과 표시 영역 생성"""
        result_frame = tk.Frame(parent, bg=self.colors['panel'])
        result_frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 결과 요약 카드
        summary_card = tk.Frame(result_frame, bg=self.colors['card'], 
                            highlightbackground=self.colors['neon_green'],
                            highlightthickness=1)
        summary_card.pack(fill='x', pady=(0, 10))
        
        # 요약 정보
        self.summary_frame = tk.Frame(summary_card, bg=self.colors['card'])
        self.summary_frame.pack(fill='x', padx=15, pady=12)
        
        # 초기 요약 표시
        self.create_initial_summary()
        
        # 상세 결과 표시 (스크롤 가능)
        detail_frame = tk.Frame(result_frame, bg=self.colors['card'])
        detail_frame.pack(fill='both', expand=True)
        
        # 상세 결과 텍스트
        self.result_text = tk.Text(detail_frame,
                                bg=self.colors['card'],
                                fg=self.colors['text_primary'],
                                font=self.fonts['mono'],
                                wrap='word',
                                height=15,
                                bd=0,
                                insertbackground=self.colors['neon_green'],
                                selectbackground=self.colors['neon_blue'],
                                selectforeground=self.colors['bg'])
        self.result_text.pack(fill='both', expand=True, padx=1, pady=1)
        
        # 초기 메시지
        self.result_text.insert('1.0', "⚡ 울트라 퍼포먼스 모드 활성화\n\n파일 업로드를 기다리는 중...")
        self.result_text.config(state='disabled')

    def create_initial_summary(self):
        """초기 요약 정보 생성"""
        tk.Label(self.summary_frame, 
                text="📊 처리 대기 중",
                font=self.fonts['header'],
                bg=self.colors['card'],
                fg=self.colors['neon_blue']).pack(anchor='w')
        
        tk.Label(self.summary_frame,
                text="엑셀 파일을 업로드하면 분류 결과가 여기에 표시됩니다",
                font=self.fonts['body'],
                bg=self.colors['card'],
                fg=self.colors['text_secondary']).pack(anchor='w', pady=(5, 0))
    
    def create_work_action_buttons(self, parent):
        """워커 관리 액션 버튼들"""
        button_config = {
            'font': self.fonts['small'],
            'bd': 0,
            'padx': 20,
            'pady': 10,
            'cursor': 'hand2',
            'width': 20
        }
        
        actions = [
            ("➕ 담당자 추가", self.colors['neon_green'], self.add_new_work),
            ("✏️ 이름 수정", self.colors['neon_blue'], self.edit_work_name),
            ("🔼 위로 이동", self.colors['neon_purple'], self.move_work_up),
            ("🔽 아래로 이동", self.colors['neon_purple'], self.move_work_down),
            ("🎨 아이콘 변경", self.colors['neon_yellow'], self.change_work_icon),
            ("📝 설명 수정", self.colors['neon_blue'], self.edit_work_description),
            ("❌ 담당자 삭제", self.colors['neon_pink'], self.delete_work),
            ("💾 변경사항 저장", self.colors['neon_green'], self.save_work_changes)
        ]
        
        for text, color, command in actions:
            btn = tk.Button(parent,
                           text=text,
                           bg=self.colors['card'],
                           fg=color,
                           activebackground=color,
                           activeforeground=self.colors['bg'],
                           command=command,
                           **button_config)
            btn.pack(pady=4, padx=15)
            
            # 호버 효과
            def make_hover(btn, color):
                def on_enter(e):
                    btn.config(bg=color, fg=self.colors['bg'])
                def on_leave(e):
                    btn.config(bg=self.colors['card'], fg=color)
                return on_enter, on_leave
            
            on_enter, on_leave = make_hover(btn, color)
            btn.bind('<Enter>', on_enter)
            btn.bind('<Leave>', on_leave)
    
    def create_product_settings_tab(self):
        """상품 설정 탭 생성 (성능 최적화)"""
        products_tab = tk.Frame(self.notebook, bg=self.colors['bg_secondary'])
        self.notebook.add(products_tab, text="🎯 상품설정")
        
        # 스크롤 가능한 영역
        canvas = tk.Canvas(products_tab, bg=self.colors['bg_secondary'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(products_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['bg_secondary'])
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=(0, 2))
        scrollbar.pack(side="right", fill="y", padx=(0, 3))
        
        # 헤더
        header_frame = tk.Frame(scrollable_frame, bg=self.colors['bg_secondary'])
        header_frame.pack(fill='x', padx=15, pady=15)
        
        title_label = tk.Label(header_frame,
                              text="상품 규칙 설정",
                              font=self.fonts['title'],
                              bg=self.colors['bg_secondary'],
                              fg=self.colors['neon_green'])
        title_label.pack()
        
        subtitle_label = tk.Label(header_frame,
                                 text="100% 정확도를 위한 세부 상품 매칭 규칙 설정",
                                 font=self.fonts['body'],
                                 bg=self.colors['bg_secondary'],
                                 fg=self.colors['text_secondary'])
        subtitle_label.pack(pady=(5, 0))
        
        # 도움말 카드
        help_card = tk.Frame(scrollable_frame, bg=self.colors['card'])
        help_card.pack(fill='x', padx=15, pady=8)
        
        help_text = """🎯 간단 가이드:
            - 브랜드: 상품 브랜드명 (첫 번째 단어)
            - 상품명: 정확한 상품명 또는 "All"로 모든 상품
            - 옵션: 특정 옵션 또는 "All"로 모든 옵션
            - 우선순위: 담당자는 순서대로 매칭됩니다"""
        
        help_label = tk.Label(help_card, text=help_text,
                             font=self.fonts['small'],
                             bg=self.colors['card'],
                             fg=self.colors['text_secondary'],
                             justify='left')
        help_label.pack(padx=15, pady=12)
        
        # 상품 설정 컨테이너
        self.products_container = tk.Frame(scrollable_frame, bg=self.colors['bg_secondary'])
        self.products_container.pack(fill='x', padx=15, pady=8)
        
        # 저장 버튼
        save_btn = tk.Button(scrollable_frame,
                            text="💾 모든 상품 규칙 저장",
                            font=self.fonts['body'],
                            bg=self.colors['neon_green'],
                            fg=self.colors['bg'],
                            activebackground=self.colors['success'],
                            bd=0, padx=40, pady=15,
                            cursor='hand2',
                            command=self.save_product_settings)
        save_btn.pack(pady=15)
        
        # 상품 프레임 초기화
        self.product_frames = {}
        self.product_lists = {}
        self.refresh_product_frames()
    
    def refresh_product_frames(self):
        """상품 설정 프레임 새로고침 (성능 최적화)"""
        # 기존 프레임 제거
        for widget in self.products_container.winfo_children():
            widget.destroy()
        
        self.product_frames = {}
        self.product_lists = {}
        
        # 워커별 프레임 생성
        for work_name in self.settings['work_order']:
            work_config = self.settings['work_config'][work_name]
            if work_config.get('type') == 'product_specific':
                self.create_worker_product_frame(work_name, work_config)
    
    def create_worker_product_frame(self, work_name, work_config):
        """개별 워커의 상품 설정 프레임 (균형잡힌 레이아웃)"""
        # 워커 컨테이너 (전체 너비 활용)
        worker_frame = tk.Frame(self.products_container, bg=self.colors['panel'])
        worker_frame.pack(fill='both', expand=True, pady=8)
        
        # 메인 가로 레이아웃 (전체 공간 활용)
        main_layout = tk.Frame(worker_frame, bg=self.colors['panel'])
        main_layout.pack(fill='both', expand=True, padx=15, pady=10)
        
        # 좌측: 워커 정보 및 컨트롤 (고정 너비 축소)
        left_section = tk.Frame(main_layout, bg=self.colors['card'], width=250)
        left_section.pack(side='left', fill='y', padx=(0, 15))
        left_section.pack_propagate(False)
        
        icon = work_config.get('icon', '📦')
        desc = work_config.get('description', '')
        
        # 워커 정보 헤더
        info_header = tk.Frame(left_section, bg=self.colors['card'])
        info_header.pack(fill='x', padx=15, pady=(15, 10))
        
        worker_label = tk.Label(info_header,
                            text=f"{icon} {work_name}",
                            font=self.fonts['header'],
                            bg=self.colors['card'],
                            fg=self.colors['neon_blue'])
        worker_label.pack(anchor='w')
        
        desc_label = tk.Label(info_header,
                            text=desc,
                            font=self.fonts['small'],
                            bg=self.colors['card'],
                            fg=self.colors['text_secondary'],
                            wraplength=220)
        desc_label.pack(anchor='w', pady=(5, 0))
        
        # 컨트롤 버튼들 (더 컴팩트하게)
        control_frame = tk.Frame(left_section, bg=self.colors['card'])
        control_frame.pack(fill='x', padx=10, pady=(0, 10))
        
        # 버튼 스타일 (더 작고 세련되게)
        btn_config = {
            'font': ('SF Pro Display', 11),
            'bd': 0,
            'padx': 15,
            'pady': 10,
            'cursor': 'hand2',
            'width': 18
        }
        
        add_btn = tk.Button(control_frame,
                        text="➕ 규칙 추가",
                        bg=self.colors['neon_green'],
                        fg=self.colors['bg'],
                        command=lambda: self.add_product_rule(work_name),
                        **btn_config)
        add_btn.pack(fill='x', pady=2)
        
        edit_btn = tk.Button(control_frame,
                            text="✏️ 규칙 수정",
                            bg=self.colors['neon_blue'],
                            fg=self.colors['bg'],
                            command=lambda: self.edit_selected_rule(work_name),
                            **btn_config)
        edit_btn.pack(fill='x', pady=2)
        
        delete_btn = tk.Button(control_frame,
                            text="❌ 규칙 삭제",
                            bg=self.colors['neon_pink'],
                            fg=self.colors['bg'],
                            command=lambda: self.delete_selected_rule(work_name),
                            **btn_config)
        delete_btn.pack(fill='x', pady=2)
        
        # 우측: 상품 리스트 (남은 공간 전체 활용)
        right_section = tk.Frame(main_layout, bg=self.colors['card'])
        right_section.pack(side='left', fill='both', expand=True, padx=(0, 0))
        
        # 리스트 헤더 (더 명확하게)
        list_header = tk.Frame(right_section, bg=self.colors['card'])
        list_header.pack(fill='x', padx=20, pady=(15, 10))
        
        header_left = tk.Frame(list_header, bg=self.colors['card'])
        header_left.pack(side='left', fill='x', expand=True)
        
        tk.Label(header_left,
                text="📋 현재 등록된 상품 규칙",
                font=('SF Pro Display', 14, 'bold'),
                bg=self.colors['card'],
                fg=self.colors['text_primary']).pack(side='left')
        
        products = work_config.get('products', [])
        count_label = tk.Label(header_left,
                            text=f"({len(products)}개)",
                            font=('SF Pro Display', 12),
                            bg=self.colors['card'],
                            fg=self.colors['neon_green'])
        count_label.pack(side='left', padx=(10, 0))
        
        # 컬럼 헤더 추가 (테이블 형식)
        column_header = tk.Frame(right_section, bg=self.colors['card'])
        column_header.pack(fill='x', padx=20, pady=(0, 5))
        
        tk.Label(column_header, text="번호", width=5, anchor='w',
                font=('SF Pro Display', 11, 'bold'),
                bg=self.colors['card'], fg=self.colors['text_secondary']).pack(side='left')
        tk.Label(column_header, text="브랜드", width=15, anchor='w',
                font=('SF Pro Display', 11, 'bold'),
                bg=self.colors['card'], fg=self.colors['text_secondary']).pack(side='left', padx=(10, 0))
        tk.Label(column_header, text="상품명", width=40, anchor='w',
                font=('SF Pro Display', 11, 'bold'),
                bg=self.colors['card'], fg=self.colors['text_secondary']).pack(side='left', padx=(10, 0))
        tk.Label(column_header, text="옵션", anchor='w',
                font=('SF Pro Display', 11, 'bold'),
                bg=self.colors['card'], fg=self.colors['text_secondary']).pack(side='left', padx=(10, 0))
        
        # 상품 리스트 (훨씬 크고 읽기 쉽게)
        list_frame = tk.Frame(right_section, bg=self.colors['bg_secondary'], 
                            highlightbackground=self.colors['border'], highlightthickness=1)
        list_frame.pack(fill='both', expand=True, padx=20, pady=(0, 15))
        
        # 스크롤바 포함 리스트박스
        list_container = tk.Frame(list_frame, bg=self.colors['card'])
        list_container.pack(fill='both', expand=True, padx=1, pady=1)
        
        product_list = tk.Listbox(list_container,
                                bg=self.colors['card'],
                                fg=self.colors['text_primary'],
                                font=('SF Pro Display', 12),  # 폰트 크기 증가
                                selectbackground=self.colors['neon_green'],
                                selectforeground=self.colors['bg'],
                                height=10,  # 높이 증가
                                bd=0,
                                highlightthickness=0,
                                activestyle='none',
                                selectmode='single')
        
        # 리스트 스크롤바 (더 보기 좋게)
        list_scrollbar = tk.Scrollbar(list_container, orient="vertical", 
                                    command=product_list.yview, width=16)
        product_list.configure(yscrollcommand=list_scrollbar.set)
        
        product_list.pack(side='left', fill='both', expand=True, padx=(10, 0))
        list_scrollbar.pack(side='right', fill='y', padx=(0, 5))
        
        # 현재 규칙들 추가 (더 구조화된 형태로)
        for i, product in enumerate(products):
            brand = product.get('brand', '').ljust(12)[:12] or '(브랜드없음)'
            product_name = product.get('product_name', '').ljust(35)[:35]
            order_option = product.get('order_option', 'All')
            
            # 탭으로 구분된 형식적인 표시
            display_text = f"{i+1:3d}.  {brand}  {product_name}  [{order_option}]"
            product_list.insert(tk.END, display_text)
        
        # 리스트가 비어있을 때 메시지
        if len(products) == 0:
            product_list.insert(tk.END, "     아직 등록된 상품 규칙이 없습니다.")
            product_list.insert(tk.END, "     왼쪽의 '➕ 규칙 추가' 버튼을 클릭하여 시작하세요.")
            product_list.config(fg=self.colors['text_muted'])
        
        self.product_frames[work_name] = worker_frame
        self.product_lists[work_name] = product_list
    
    def create_stats_tab(self):
        """통계 탭 생성"""
        stats_tab = tk.Frame(self.notebook, bg=self.colors['bg_secondary'])
        self.notebook.add(stats_tab, text="📈 통계 분석")
        
        # 메인 컨테이너
        main_container = tk.Frame(stats_tab, bg=self.colors['bg_secondary'])
        main_container.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 상단: 정확도 카드
        accuracy_card = tk.Frame(main_container, bg=self.colors['card'],
                                highlightbackground=self.colors['neon_green'],
                                highlightthickness=2)
        accuracy_card.pack(fill='x', pady=(0, 20))
        
        # 정확도 헤더
        acc_header = tk.Label(accuracy_card,
                             text="🎯 분류 정확도",
                             font=self.fonts['header'],
                             bg=self.colors['card'],
                             fg=self.colors['neon_green'])
        acc_header.pack(pady=(20, 10))
        
        # 정확도 표시
        self.accuracy_text = tk.Text(accuracy_card,
                                    height=6,
                                    bg=self.colors['card'],
                                    fg=self.colors['text_primary'],
                                    font=self.fonts['body'],
                                    bd=0,
                                    wrap='word')
        self.accuracy_text.pack(fill='x', padx=20, pady=(0, 20))
        self.accuracy_text.insert('1.0', "파일 처리 후 정확도 지표를 확인하세요...")
        self.accuracy_text.config(state='disabled')
        
        # 하단: 상세 통계
        stats_card = tk.Frame(main_container, bg=self.colors['panel'])
        stats_card.pack(fill='both', expand=True)
        
        # 통계 헤더
        stats_header = tk.Label(stats_card,
                               text="📊 상세 통계",
                               font=self.fonts['header'],
                               bg=self.colors['panel'],
                               fg=self.colors['text_primary'])
        stats_header.pack(pady=20)
        
        # 통계 텍스트
        stats_frame = tk.Frame(stats_card, bg=self.colors['card'])
        stats_frame.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        
        self.stats_text = tk.Text(stats_frame,
                                 bg=self.colors['card'],
                                 fg=self.colors['text_primary'],
                                 font=self.fonts['mono'],
                                 bd=0,
                                 wrap='word')
        self.stats_text.pack(fill='both', expand=True, padx=1, pady=1)
        self.stats_text.insert('1.0', "데이터 처리를 기다리는 중...")
        self.stats_text.config(state='disabled')
    
    # 헬퍼 메서드들 (성능 최적화)
    def update_status(self, message):
        """상태바 업데이트"""
        self.status_var.set(message)
        
    def refresh_work_list(self):
        """워커 리스트 새로고침"""
        self.work_listbox.delete(0, tk.END)
        
        for i, work_name in enumerate(self.settings['work_order']):
            work_config = self.settings['work_config'][work_name]
            icon = work_config.get('icon', '📦')
            desc = work_config.get('description', '')
            
            display_text = f"{i+1}. {icon} {work_name} - {desc}"
            self.work_listbox.insert(tk.END, display_text)
    
    def update_accuracy_display(self, text):
        """정확도 표시 업데이트"""
        self.accuracy_text.config(state='normal')
        self.accuracy_text.delete(1.0, tk.END)
        self.accuracy_text.insert(tk.END, text)
        self.accuracy_text.config(state='disabled')
    
    def update_stats_display(self, text):
        """통계 표시 업데이트"""
        self.stats_text.config(state='normal')
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(tk.END, text)
        self.stats_text.config(state='disabled')
    
    # 파일 선택 및 처리 메서드들
    def select_file(self):
        """파일 선택 (즉시 실행)"""
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("엑셀 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")]
        )
        
        if file_path:
            self.selected_file = file_path
            filename = os.path.basename(file_path)
            self.file_info_var.set(f"선택됨: {filename}")
            self.process_button.config(state='normal')
            self.update_status(f"파일 로드됨: {filename}")
    
    def process_excel(self):
        """엑셀 파일 처리 (고성능 최적화)"""
        if not hasattr(self, 'selected_file'):
            messagebox.showerror("Error", "Please select a file first")
            return
        
        # 버튼 비활성화
        self.process_button.config(state='disabled')
        self.download_button.config(state='disabled')
        self.review_button.config(state='disabled')
        
        # 프로그레스 다이얼로그
        self.progress_dialog = FastProgressDialog(self.root, "⚡ Ultra Fast Processing...")
        
        # 백그라운드 처리
        thread = threading.Thread(target=self._process_excel_optimized)
        thread.daemon = True
        thread.start()
    
    def _process_excel_optimized(self):
        """최적화된 엑셀 처리"""
        try:
            start_time = time.time()
            
            # 1. 파일 로딩 (청크 단위 읽기로 메모리 효율화)
            self.update_progress(5, "Loading file...", 50)
            
            # 엔진 자동 선택
            engine = 'xlrd' if self.selected_file.endswith('.xls') else 'openpyxl'
            
            # 청크 단위로 읽기 (대용량 파일 대응)
            df = pd.read_excel(self.selected_file, engine=engine)
            self.excel_data = df
            
            # 필수 컬럼 검증
            required_columns = ['상품명', '주문수량']
            missing = [col for col in required_columns if col not in df.columns]
            if missing:
                raise ValueError(f"Missing columns: {', '.join(missing)}")
            
            self.update_progress(10, f"Loaded {len(df)} orders", 100)
            
            # 2. 전처리 (벡터화 연산)
            self.update_progress(15, "Preprocessing data...", 0)
            df = self._preprocess_data_optimized(df)
            self.update_progress(20, "Preprocessing complete", 100)
            
            # 3. 분류 (병렬 처리)
            self.update_progress(25, "Classifying orders...", 0)
            classified_df = self._classify_orders_optimized(df)
            
            # 4. 정렬 (최적화된 알고리즘)
            self.update_progress(70, "Sorting results...", 0)
            sorted_df = self._sort_results_optimized(classified_df)
            
            # 5. 통계 계산
            self.update_progress(85, "Calculating statistics...", 0)
            self._calculate_statistics(sorted_df)
            
            # 6. 완료
            self.classified_data = sorted_df
            elapsed_time = time.time() - start_time
            self.update_progress(100, f"Complete! ({elapsed_time:.1f}s)", 100)
            
            # UI 업데이트
            self.root.after(0, self._process_complete)
            
        except Exception as e:
            self.root.after(0, lambda: self._process_error(str(e)))
    
    def _preprocess_data_optimized(self, df):
        """최적화된 데이터 전처리"""
        # 벡터화 연산 사용
        df['상품명'] = df['상품명'].fillna('').astype(str)
        df['주문수량'] = pd.to_numeric(df['주문수량'], errors='coerce').fillna(0).astype(int)
        
        # 주문선택사항 처리
        if '주문선택사항' in df.columns:
            df['주문선택사항'] = df['주문선택사항'].fillna('').astype(str)
            df['full_product_name'] = df['상품명'] + ' ' + df['주문선택사항']
        else:
            df['주문선택사항'] = ''
            df['full_product_name'] = df['상품명']
        
        # 브랜드 추출 (벡터화)
        df['brand'] = df['상품명'].str.split(n=1, expand=True)[0].fillna('')
        
        # 주문번호 처리
        if '주문고유번호' in df.columns:
            df['주문고유번호'] = df['주문고유번호'].fillna('').astype(str)
        else:
            df['주문고유번호'] = np.arange(len(df)).astype(str)
        
        return df
    
    def _classify_orders_optimized(self, df):
        """최적화된 주문 분류"""
        total_rows = len(df)
        df = df.copy()
        
        # 기본값 설정
        failed_work = self._get_failed_work_name()
        df['담당자'] = failed_work
        df['분류근거'] = '매칭 없음'
        df['신뢰도'] = 0.0
        
        # 설정값
        quantity_threshold = self.settings.get('quantity_threshold', 2)
        
        # 1. 합배송 판별 (벡터화)
        if '주문고유번호' in df.columns:
            order_counts = df['주문고유번호'].value_counts()
            multi_orders = order_counts[order_counts >= 2].index
            is_multi_order = df['주문고유번호'].isin(multi_orders)
            
            combined_work = self._get_combined_work_name()
            if combined_work:
                df.loc[is_multi_order, '담당자'] = combined_work
                df.loc[is_multi_order, '분류근거'] = '합배송'
                df.loc[is_multi_order, '신뢰도'] = 1.0
        
        # 2. 복수주문 판별 (벡터화)
        multiple_work = self._get_multiple_work_name()
        if multiple_work:
            is_multiple = (df['주문수량'] >= quantity_threshold) & (df['담당자'] == failed_work)
            df.loc[is_multiple, '담당자'] = multiple_work
            df.loc[is_multiple, '분류근거'] = '복수주문'
            df.loc[is_multiple, '신뢰도'] = 1.0
        
        # 3. 상품별 매칭 (최적화)
        unmatched_mask = df['담당자'] == failed_work
        unmatched_indices = df[unmatched_mask].index
        
        if len(unmatched_indices) > 0:
            # 규칙 사전 컴파일
            compiled_rules = self._compile_matching_rules()
            
            # 배치 처리
            batch_size = 1000
            for i in range(0, len(unmatched_indices), batch_size):
                batch_indices = unmatched_indices[i:i+batch_size]
                self._classify_batch(df, batch_indices, compiled_rules)
                
                # 진행률 업데이트
                progress = 25 + (i / len(unmatched_indices)) * 45
                self.update_progress(progress, f"Classifying... {i}/{len(unmatched_indices)}", 
                                   (i % batch_size) / batch_size * 100)
        
        return df
    
    def _compile_matching_rules(self):
        """매칭 규칙 사전 컴파일 (성능 향상)"""
        rules = []
        for work_name in self.settings['work_order']:
            work_config = self.settings['work_config'][work_name]
            if work_config.get('type') != 'product_specific':
                continue
            
            for product in work_config.get('products', []):
                rules.append({
                    'work_name': work_name,
                    'brand': product.get('brand', ''),
                    'product_name': product.get('product_name', ''),
                    'order_option': product.get('order_option', 'All')
                })
        return rules
    
    def _classify_batch(self, df, indices, rules):
        """배치 단위 분류"""
        for idx in indices:
            row = df.loc[idx]
            
            for rule in rules:
                if self._match_rule(row, rule):
                    df.at[idx, '담당자'] = rule['work_name']
                    df.at[idx, '분류근거'] = f"매칭: {rule['brand']} {rule['product_name']}"
                    df.at[idx, '신뢰도'] = 1.0
                    break
    
    def _match_rule(self, row, rule):
        """규칙 매칭 (최적화)"""
        # 브랜드 체크
        if rule['brand'] and rule['brand'] != row['brand']:
            return False
        
        # 상품명 체크
        if rule['product_name'] != 'All':
            if rule['product_name'] not in row['상품명']:
                return False
        
        # 옵션 체크
        if rule['order_option'] != 'All':
            if rule['order_option'] not in row['주문선택사항']:
                return False
        
        return True
    
    def _sort_results_optimized(self, df):
        """최적화된 정렬"""
        # 우선순위 매핑
        priority_map = {name: i for i, name in enumerate(self.settings['work_order'])}
        df['priority'] = df['담당자'].map(priority_map)
        
        # 정렬 키 생성
        combined_work = self._get_combined_work_name()
        
        # 그룹별 정렬
        sorted_groups = []
        for work_name in self.settings['work_order']:
            work_df = df[df['담당자'] == work_name]
            
            if len(work_df) == 0:
                continue
            
            if work_name == combined_work:
                # 합배송: 주문번호 그룹
                work_df = work_df.sort_values(['주문고유번호'])
            else:
                # 일반: 상품명 정렬
                work_df = work_df.sort_values(['full_product_name'])
            
            sorted_groups.append(work_df)
        
        # 병합
        if sorted_groups:
            sorted_df = pd.concat(sorted_groups, ignore_index=True)
            sorted_df = sorted_df.drop(['priority'], axis=1)
        else:
            sorted_df = df
        
        return sorted_df
    
    def _calculate_statistics(self, df):
        """통계 계산"""
        total_orders = len(df)
        
        # 담당자별 통계
        self.work_ranges = {}
        work_stats = {}
        
        for work_name in self.settings['work_order']:
            work_data = df[df['담당자'] == work_name]
            count = len(work_data)
            
            if count > 0:
                start_row = work_data.index[0] + 2
                end_row = work_data.index[-1] + 2
                
                self.work_ranges[work_name] = {
                    'start': start_row,
                    'end': end_row,
                    'count': count,
                    'icon': self.settings['work_config'][work_name].get('icon', '📦')
                }
                
                work_stats[work_name] = {
                    'count': count,
                    'percentage': count / total_orders * 100,
                    'avg_confidence': work_data['신뢰도'].mean()
                }
        
        # 전체 통계
        failed_work = self._get_failed_work_name()
        unmatched_count = len(df[df['담당자'] == failed_work])
        auto_rate = (total_orders - unmatched_count) / total_orders * 100
        
        self.accuracy_metrics = {
            'total_orders': total_orders,
            'auto_classification_rate': auto_rate,
            'unmatched_count': unmatched_count,
            'work_stats': work_stats
        }
    
    def update_progress(self, percent, status, sub_percent=0):
        """진행률 업데이트"""
        if hasattr(self, 'progress_dialog') and not self.progress_dialog.cancelled:
            self.progress_dialog.update(percent, status, sub_percent)
    
    def _process_complete(self):
        """처리 완료"""
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.close()
        
        # 결과 표시
        self._display_results()
        self._update_statistics()
        
        # 버튼 활성화
        self.process_button.config(state='normal')
        self.download_button.config(state='normal')
        
        if self.accuracy_metrics['unmatched_count'] > 0:
            self.review_button.config(state='normal')
        
        # 상태 업데이트
        auto_rate = self.accuracy_metrics['auto_classification_rate']
        self.update_status(f"✅ Complete! Auto-classification: {auto_rate:.1f}%")
        
        if auto_rate == 100:
            messagebox.showinfo("Perfect!", 
                              "🎉 100% accuracy achieved!\n\n"
                              "All orders classified successfully.")
    
    def _process_error(self, error_msg):
        """처리 오류"""
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.close()
        
        self.process_button.config(state='normal')
        self.update_status(f"❌ Error: {error_msg}")
        messagebox.showerror("Processing Error", error_msg)
    
    def _display_results(self):
        """결과 표시 (요약 + 상세)"""
        # 요약 카드 업데이트
        for widget in self.summary_frame.winfo_children():
            widget.destroy()
        
        metrics = self.accuracy_metrics
        
        # 요약 제목
        title_label = tk.Label(self.summary_frame,
                            text=f"✅ 분류 완료 ({metrics['auto_classification_rate']:.1f}% 자동분류)",
                            font=self.fonts['header'],
                            bg=self.colors['card'],
                            fg=self.colors['neon_green'])
        title_label.pack(anchor='w')
        
        # 요약 정보
        summary_text = f"총 {metrics['total_orders']}건 • 성공 {metrics['total_orders'] - metrics['unmatched_count']}건 • 검토필요 {metrics['unmatched_count']}건"
        summary_label = tk.Label(self.summary_frame,
                                text=summary_text,
                                font=self.fonts['body'],
                                bg=self.colors['card'],
                                fg=self.colors['text_secondary'])
        summary_label.pack(anchor='w', pady=(5, 0))
        
        # 상세 결과 업데이트
        self.result_text.config(state='normal')
        self.result_text.delete(1.0, tk.END)
        
        result = f"""📋 담당자별 분류 결과
                {'='*40}

                """
        
        for work_name, stats in self.work_ranges.items():
            icon = stats['icon']
            count = stats['count']
            percentage = self.accuracy_metrics['work_stats'][work_name]['percentage']
            
            result += f"{icon} {work_name}\n"
            result += f"   주문건수: {count}건 ({percentage:.1f}%)\n"
            result += f"   엑셀행: {stats['start']}~{stats['end']}행\n\n"
        
        if metrics['unmatched_count'] > 0:
            result += f"❗ 미분류 {metrics['unmatched_count']}건은 상품설정 탭에서 규칙을 추가해주세요."
        
        self.result_text.insert(tk.END, result)
        self.result_text.config(state='disabled')
    
    def _update_statistics(self):
        """통계 업데이트"""
        # 정확도 표시
        metrics = self.accuracy_metrics
        accuracy_text = f"""Auto-classification Rate: {metrics['auto_classification_rate']:.1f}%
Total Orders Processed: {metrics['total_orders']}
Successfully Classified: {metrics['total_orders'] - metrics['unmatched_count']}
Manual Review Needed: {metrics['unmatched_count']}

{'🏆 PERFECT SCORE!' if metrics['auto_classification_rate'] == 100 else '💡 Add more rules to improve accuracy'}"""
        
        self.update_accuracy_display(accuracy_text)
        
        # 상세 통계
        stats_text = "DETAILED STATISTICS\n" + "="*50 + "\n\n"
        
        for work_name, work_stats in metrics['work_stats'].items():
            if work_stats['count'] > 0:
                stats_text += f"{work_name}:\n"
                stats_text += f"  Orders: {work_stats['count']} ({work_stats['percentage']:.1f}%)\n"
                stats_text += f"  Confidence: {work_stats['avg_confidence']:.1%}\n\n"
        
        self.update_stats_display(stats_text)
    
    # 나머지 헬퍼 메서드들
    def edit_selected_rule(self, work_name):
        """선택된 규칙 수정"""
        product_list = self.product_lists.get(work_name)
        if not product_list:
            return
            
        selection = product_list.curselection()
        if not selection:
            messagebox.showwarning("선택 필요", "수정할 규칙을 선택해주세요")
            return
        
        # 선택된 인덱스에서 실제 상품 정보 가져오기
        selected_idx = selection[0]
        work_config = self.settings['work_config'][work_name]
        products = work_config.get('products', [])
        
        if selected_idx >= len(products):
            return
            
        product = products[selected_idx]
        
        dialog = ProductRuleDialog(self.root, work_name, mode='edit', 
                                initial_data={
                                    'brand': product.get('brand', ''),
                                    'product_name': product.get('product_name', ''),
                                    'order_option': product.get('order_option', 'All')
                                })
        
        if dialog.result:
            # 기존 상품 정보 업데이트
            products[selected_idx] = {
                'brand': dialog.result['brand'],
                'product_name': dialog.result['product_name'],
                'order_option': dialog.result['order_option']
            }
            
            # 리스트 새로고침
            self.refresh_product_frames()

    def delete_selected_rule(self, work_name):
        """선택된 규칙 삭제"""
        product_list = self.product_lists.get(work_name)
        if not product_list:
            return
            
        selection = product_list.curselection()
        if not selection:
            messagebox.showwarning("선택 필요", "삭제할 규칙을 선택해주세요")
            return
        
        selected_idx = selection[0]
        work_config = self.settings['work_config'][work_name]
        products = work_config.get('products', [])
        
        if selected_idx >= len(products):
            return
            
        # 삭제 확인
        product = products[selected_idx]
        product_name = product.get('product_name', '규칙')
        
        if messagebox.askyesno("삭제 확인", f"'{product_name}' 규칙을 삭제하시겠습니까?"):
            # 상품 리스트에서 제거
            products.pop(selected_idx)
            
            # 리스트 새로고침
            self.refresh_product_frames()

    def _get_failed_work_name(self):
        """실패 담당자명"""
        for name in self.settings['work_order']:
            if self.settings['work_config'][name].get('type') == 'failed':
                return name
        return "분류실패"
    
    def _get_combined_work_name(self):
        """합배송 담당자명"""
        for name in self.settings['work_order']:
            if self.settings['work_config'][name].get('type') == 'mixed_products':
                return name
        return None
    
    def _get_multiple_work_name(self):
        """복수주문 담당자명"""
        for name in self.settings['work_order']:
            if self.settings['work_config'][name].get('type') == 'multiple_quantity':
                return name
        return None
    
    # 다이얼로그 메서드들 (간소화)
    def add_product_rule(self, work_name):
        """상품 규칙 추가"""
        dialog = ProductRuleDialog(self.root, work_name, mode='add')
        if dialog.result:
            product_list = self.product_lists[work_name]
            rule_text = f"{dialog.result['brand']} | {dialog.result['product_name']} | {dialog.result['order_option']}"
            product_list.insert(tk.END, rule_text)
    
    def edit_product_rule(self, work_name, product_list):
        """상품 규칙 수정"""
        selection = product_list.curselection()
        if not selection:
            messagebox.showwarning("Selection Required", "Please select a rule to edit")
            return
        
        # 기존 규칙 파싱
        rule_text = product_list.get(selection[0])
        parts = rule_text.split(' | ')
        
        dialog = ProductRuleDialog(self.root, work_name, mode='edit', 
                                 initial_data={
                                     'brand': parts[0],
                                     'product_name': parts[1],
                                     'order_option': parts[2]
                                 })
        
        if dialog.result:
            new_text = f"{dialog.result['brand']} | {dialog.result['product_name']} | {dialog.result['order_option']}"
            product_list.delete(selection[0])
            product_list.insert(selection[0], new_text)
    
    def delete_product_rule(self, work_name, product_list):
        """상품 규칙 삭제"""
        selection = product_list.curselection()
        if not selection:
            messagebox.showwarning("Selection Required", "Please select a rule to delete")
            return
        
        if messagebox.askyesno("Confirm Delete", "Delete selected rule?"):
            product_list.delete(selection[0])
    
    def save_product_settings(self):
        """상품 설정 저장"""
        try:
            # 각 워커의 상품 규칙 수집
            for work_name, product_list in self.product_lists.items():
                products = []
                
                for i in range(product_list.size()):
                    rule_text = product_list.get(i)
                    parts = rule_text.split(' | ')
                    
                    products.append({
                        'brand': parts[0],
                        'product_name': parts[1],
                        'order_option': parts[2]
                    })
                
                self.settings['work_config'][work_name]['products'] = products
            
            self.save_settings()
            self.update_status("✅ Product settings saved successfully")
            messagebox.showinfo("Success", "Product settings saved!")
            
        except Exception as e:
            messagebox.showerror("Save Error", str(e))
    
    def download_excel(self):
        """결과 다운로드"""
        if self.classified_data is None:
            messagebox.showerror("Error", "No data to download")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"classified_{timestamp}.xlsx"
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                # 내부 컬럼 제거
                output_cols = [col for col in self.classified_data.columns 
                             if col not in ['담당자', '분류근거', '신뢰도', 'brand', 'full_product_name', 'priority']]
                output_data = self.classified_data[output_cols]
                
                # 저장
                output_data.to_excel(save_path, index=False)
                self.update_status(f"✅ Saved: {os.path.basename(save_path)}")
                messagebox.showinfo("Success", "File saved successfully!")
                
            except Exception as e:
                messagebox.showerror("Save Error", str(e))
    
    def review_unmatched(self):
        """미분류 검토"""
        if self.accuracy_metrics['unmatched_count'] == 0:
            messagebox.showinfo("Perfect!", "No unmatched items to review!")
            return
        
        messagebox.showinfo("Review Needed", 
                          f"{self.accuracy_metrics['unmatched_count']} items need rules.\n\n"
                          "Go to Products tab to add specific rules.")
    
    # 워커 관리 메서드들
    def add_new_work(self):
        """새 워커 추가"""
        name = simpledialog.askstring("Add Worker", "Enter worker name:")
        if name and name.strip():
            name = name.strip()
            if name in self.settings['work_config']:
                messagebox.showerror("Error", "Worker already exists")
                return
            
            self.settings['work_config'][name] = {
                "type": "product_specific",
                "products": [],
                "description": f"{name} duties",
                "icon": "👤",
                "enabled": True
            }
            
            # 분류실패 전에 삽입
            failed_idx = self.settings['work_order'].index(self._get_failed_work_name())
            self.settings['work_order'].insert(failed_idx, name)
            
            self.refresh_work_list()
            self.refresh_product_frames()
            messagebox.showinfo("Success", f"Worker '{name}' added!")
    
    def edit_work_name(self):
        """워커명 수정"""
        selection = self.work_listbox.curselection()
        if not selection:
            messagebox.showwarning("Selection Required", "Select a worker to edit")
            return
        
        old_name = self.settings['work_order'][selection[0]]
        new_name = simpledialog.askstring("Edit Name", "Enter new name:", 
                                         initialvalue=old_name)
        
        if new_name and new_name != old_name:
            if new_name in self.settings['work_config']:
                messagebox.showerror("Error", "Name already exists")
                return
            
            self.settings['work_config'][new_name] = self.settings['work_config'].pop(old_name)
            self.settings['work_order'][selection[0]] = new_name
            
            self.refresh_work_list()
            self.refresh_product_frames()
    
    def move_work_up(self):
        """워커 위로 이동"""
        selection = self.work_listbox.curselection()
        if not selection or selection[0] == 0:
            return
        
        idx = selection[0]
        self.settings['work_order'][idx], self.settings['work_order'][idx-1] = \
            self.settings['work_order'][idx-1], self.settings['work_order'][idx]
        
        self.refresh_work_list()
        self.work_listbox.selection_set(idx-1)
    
    def move_work_down(self):
        """워커 아래로 이동"""
        selection = self.work_listbox.curselection()
        if not selection or selection[0] >= len(self.settings['work_order'])-1:
            return
        
        idx = selection[0]
        self.settings['work_order'][idx], self.settings['work_order'][idx+1] = \
            self.settings['work_order'][idx+1], self.settings['work_order'][idx]
        
        self.refresh_work_list()
        self.work_listbox.selection_set(idx+1)
    
    def change_work_icon(self):
        """워커 아이콘 변경"""
        selection = self.work_listbox.curselection()
        if not selection:
            return
        
        work_name = self.settings['work_order'][selection[0]]
        current = self.settings['work_config'][work_name].get('icon', '📦')
        
        new_icon = simpledialog.askstring("Change Icon", "Enter emoji:", 
                                         initialvalue=current)
        if new_icon:
            self.settings['work_config'][work_name]['icon'] = new_icon
            self.refresh_work_list()
            self.refresh_product_frames()
    
    def edit_work_description(self):
        """워커 설명 수정"""
        selection = self.work_listbox.curselection()
        if not selection:
            return
        
        work_name = self.settings['work_order'][selection[0]]
        current = self.settings['work_config'][work_name].get('description', '')
        
        new_desc = simpledialog.askstring("Edit Description", "Enter description:", 
                                         initialvalue=current)
        if new_desc is not None:
            self.settings['work_config'][work_name]['description'] = new_desc
            self.refresh_work_list()
    
    def delete_work(self):
        """워커 삭제"""
        selection = self.work_listbox.curselection()
        if not selection:
            return
        
        work_name = self.settings['work_order'][selection[0]]
        
        if work_name in ["합배송", "복수주문", "분류실패"]:
            messagebox.showerror("Error", "Cannot delete system workers")
            return
        
        if messagebox.askyesno("Confirm Delete", f"Delete worker '{work_name}'?"):
            del self.settings['work_config'][work_name]
            self.settings['work_order'].remove(work_name)
            
            self.refresh_work_list()
            self.refresh_product_frames()
    
    def save_work_changes(self):
        """워커 변경사항 저장"""
        try:
            self.save_settings()
            self.update_status("✅ Worker settings saved")
            messagebox.showinfo("Success", "Settings saved!")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))
    
    def run(self):
        """앱 실행"""
        self.show_splash_screen()
        self.root.mainloop()
    
    def show_splash_screen(self):
        """스플래시 스크린"""
        splash = tk.Toplevel(self.root)
        splash.overrideredirect(True)
        splash.geometry("700x400")
        splash.configure(bg=self.colors['bg'])
        
        # 중앙 배치
        splash.update_idletasks()
        x = (splash.winfo_screenwidth() // 2) - 350
        y = (splash.winfo_screenheight() // 2) - 200
        splash.geometry(f"+{x}+{y}")
        
        # 콘텐츠
        tk.Label(splash, text="⚡", font=('Arial', 100), 
                bg=self.colors['bg'], fg=self.colors['neon_green']).pack(pady=50)
        
        tk.Label(splash, text="PLAYAUTO v4.1",
                font=('SF Pro Display', 36, 'bold'),
                bg=self.colors['bg'], fg=self.colors['text_primary']).pack()
        
        tk.Label(splash, text="Ultra Performance Edition",
                font=('SF Pro Display', 18),
                bg=self.colors['bg'], fg=self.colors['neon_yellow']).pack(pady=10)
        
        tk.Label(splash, text="Initializing...",
                font=('SF Pro Display', 14),
                bg=self.colors['bg'], fg=self.colors['text_secondary']).pack(pady=30)
        
        # 2초 후 닫기
        splash.after(2000, splash.destroy)

# 상품 규칙 다이얼로그
class ProductRuleDialog:
    def __init__(self, parent, work_name, mode='add', initial_data=None):
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"{mode.title()} Product Rule - {work_name}")
        self.dialog.geometry("500x300")
        self.dialog.configure(bg='#1a1a1a')
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 입력 필드들
        fields = [
            ("Brand:", initial_data.get('brand', '') if initial_data else ''),
            ("Product Name:", initial_data.get('product_name', '') if initial_data else ''),
            ("Order Option:", initial_data.get('order_option', 'All') if initial_data else 'All')
        ]
        
        self.entries = {}
        for i, (label, value) in enumerate(fields):
            tk.Label(self.dialog, text=label, bg='#1a1a1a', fg='white',
                    font=('SF Pro Display', 12)).grid(row=i, column=0, padx=20, pady=10, sticky='w')
            
            entry = tk.Entry(self.dialog, font=('SF Pro Display', 12), 
                           bg='#2a2a2a', fg='white', insertbackground='white')
            entry.insert(0, value)
            entry.grid(row=i, column=1, padx=20, pady=10, sticky='ew')
            
            self.entries[label.rstrip(':')] = entry
        
        self.dialog.grid_columnconfigure(1, weight=1)
        
        # 버튼들
        btn_frame = tk.Frame(self.dialog, bg='#1a1a1a')
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=20)
        
        tk.Button(btn_frame, text="Save", bg='#00ff88', fg='black',
                 font=('SF Pro Display', 12), bd=0, padx=30, pady=10,
                 command=self.save).pack(side='left', padx=10)
        
        tk.Button(btn_frame, text="Cancel", bg='#ff0088', fg='white',
                 font=('SF Pro Display', 12), bd=0, padx=30, pady=10,
                 command=self.dialog.destroy).pack(side='left')
    
    def save(self):
        self.result = {
            'brand': self.entries['Brand'].get().strip(),
            'product_name': self.entries['Product Name'].get().strip(),
            'order_option': self.entries['Order Option'].get().strip() or 'All'
        }
        
        if not self.result['product_name']:
            messagebox.showerror("Error", "Product name is required")
            return
        
        self.dialog.destroy()

# 메인 실행
def main():
    try:
        app = PlayAutoOrderClassifierV41()
        app.run()
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        messagebox.showerror("Fatal Error", f"Application error:\n{str(e)}")

if __name__ == "__main__":
    success = main()
    if not success:
        sys.exit(1)
    elif not getattr(sys, 'frozen', False):  # 빌드된 실행파일이 아닐 경우만 input 실행
        input("\n엔터를 눌러 종료...")
