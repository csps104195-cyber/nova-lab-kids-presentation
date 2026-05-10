#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NOVA LAB Kids Presentation Generator
Automatically generates a professional PowerPoint PPTX for product launch

Author: NOVA LAB Team
Date: 2026-05-10
Usage: python generate_pptx.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# Color Palette
COLORS = {
    'primary_blue': RGBColor(20, 40, 160),
    'secondary_blue': RGBColor(37, 57, 193),
    'sky_blue': RGBColor(0, 176, 239),
    'danger_red': RGBColor(255, 59, 48),
    'warning_yellow': RGBColor(255, 149, 0),
    'safe_green': RGBColor(52, 199, 89),
    'dark_bg': RGBColor(0, 0, 0),
    'white': RGBColor(255, 255, 255),
    'light_gray': RGBColor(248, 248, 248),
    'medium_gray': RGBColor(170, 170, 170),
    'dark_gray': RGBColor(102, 102, 102),
}

def create_presentation():
    """Create PowerPoint presentation"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    return prs

def add_title_slide(prs):
    """Slide 1: Title Page"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLORS['primary_blue']
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    p = title_frame.paragraphs[0]
    p.text = "暖口安心"
    p.font.size = Pt(66)
    p.font.bold = True
    p.font.color.rgb = COLORS['sky_blue']
    p.alignment = PP_ALIGN.CENTER
    
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9), Inches(0.8))
    subtitle_frame = subtitle_box.text_frame
    p = subtitle_frame.paragraphs[0]
    p.text = "Keep Every Bite Safe! 🍽️"
    p.font.size = Pt(32)
    p.font.color.rgb = COLORS['white']
    p.alignment = PP_ALIGN.CENTER
    
    tag_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.1), Inches(9), Inches(0.6))
    tag_frame = tag_box.text_frame
    p = tag_frame.paragraphs[0]
    p.text = "A Smart Spoon for Healthy Kids"
    p.font.size = Pt(18)
    p.font.color.rgb = COLORS['white']
    p.alignment = PP_ALIGN.CENTER
    
    footer_box = slide.shapes.add_textbox(Inches(0.5), Inches(5), Inches(9), Inches(0.4))
    footer_frame = footer_box.text_frame
    p = footer_frame.paragraphs[0]
    p.text = "NOVA LAB Team 2026 | Samsung Solve for Tomorrow"
    p.font.size = Pt(10)
    p.font.color.rgb = COLORS['medium_gray']
    p.alignment = PP_ALIGN.CENTER

def add_problem_slide(prs):
    """Slide 2: Problem Introduction"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLORS['light_gray']
    
    left_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.2), Inches(5.625))
    left_bar.fill.solid()
    left_bar.fill.fore_color.rgb = COLORS['danger_red']
    left_bar.line.color.rgb = COLORS['danger_red']
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = "為什麼需要暖口安心？"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = COLORS['primary_blue']
    
    quote_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(8), Inches(0.7))
    quote_frame = quote_box.text_frame
    p = quote_frame.paragraphs[0]
    p.text = "媽媽，這個太燙了！🔥"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = COLORS['danger_red']
    
    stats = [
        ("危險溫度", "65°C", "國際癌症研究機構認定食道癌風險臨界溫度"),
        ("年度死亡人數", "1,980", "台灣食道癌死亡案例（2022年）"),
        ("全球新診斷病例", "572K+", "2018年全球食道癌新確診病例數")
    ]
    
    y_pos = 2.1
    for label, number, description in stats:
        label_box = slide.shapes.add_textbox(Inches(0.8), Inches(y_pos), Inches(8), Inches(0.25))
        label_frame = label_box.text_frame
        p = label_frame.paragraphs[0]
        p.text = label
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = COLORS['primary_blue']
        
        num_box = slide.shapes.add_textbox(Inches(0.8), Inches(y_pos + 0.25), Inches(8), Inches(0.4))
        num_frame = num_box.text_frame
        p = num_frame.paragraphs[0]
        p.text = number
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = COLORS['danger_red']
        
        desc_box = slide.shapes.add_textbox(Inches(0.8), Inches(y_pos + 0.65), Inches(8), Inches(0.35))
        desc_frame = desc_box.text_frame
        desc_frame.word_wrap = True
        p = desc_frame.paragraphs[0]
        p.text = description
        p.font.size = Pt(10)
        p.font.color.rgb = COLORS['dark_gray']
        
        y_pos += 0.95
    
    cta_box = slide.shapes.add_textbox(Inches(0.8), Inches(4.8), Inches(8), Inches(0.5))
    cta_frame = cta_box.text_frame
    p = cta_frame.paragraphs[0]
    p.text = "✅ 但我們可以做得更好！"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLORS['safe_green']

def add_solution_slide(prs):
    """Slide 3: Solution Overview"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLORS['white']
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = "認識暖口安心"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = COLORS['primary_blue']
    
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.95), Inches(9), Inches(0.3))
    subtitle_frame = subtitle_box.text_frame
    p = subtitle_frame.paragraphs[0]
    p.text = "解決方案概述 — 三大核心功能"
    p.font.size = Pt(14)
    p.font.color.rgb = COLORS['dark_gray']
    
    features = [
        ("🌡️", "溫度感測", "知道食物有多熱", COLORS['sky_blue']),
        ("🔊", "聲音提醒", "聽到安全信號", COLORS['warning_yellow']),
        ("💡", "彩色燈號", "看懂安不安全", COLORS['safe_green'])
    ]
    
    x_positions = [0.8, 3.6, 6.4]
    for i, (emoji, title, desc, color) in enumerate(features):
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 
                                     Inches(x_positions[i]), Inches(1.8), 
                                     Inches(2.5), Inches(3))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(240, 250, 255)
        box.line.color.rgb = color
        box.line.width = Pt(2)
        
        emoji_box = slide.shapes.add_textbox(Inches(x_positions[i]), Inches(2), 
                                             Inches(2.5), Inches(0.5))
        emoji_frame = emoji_box.text_frame
        p = emoji_frame.paragraphs[0]
        p.text = emoji
        p.font.size = Pt(40)
        p.alignment = PP_ALIGN.CENTER
        
        title_box = slide.shapes.add_textbox(Inches(x_positions[i]), Inches(2.6), 
                                             Inches(2.5), Inches(0.5))
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        p = title_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = color
        p.alignment = PP_ALIGN.CENTER
        
        desc_box = slide.shapes.add_textbox(Inches(x_positions[i]), Inches(3.3), 
                                            Inches(2.5), Inches(1.2))
        desc_frame = desc_box.text_frame
        desc_frame.word_wrap = True
        p = desc_frame.paragraphs[0]
        p.text = desc
        p.font.size = Pt(11)
        p.font.color.rgb = COLORS['dark_gray']
        p.alignment = PP_ALIGN.CENTER

def add_tech_slide_1(prs):
    """Slide 4: Core Tech - Cooling & Sensing"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLORS['light_gray']
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = "核心技術 (1/2) — 致冷 + 感測"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLORS['primary_blue']
    
    heading_box = slide.shapes.add_textbox(Inches(5.5), Inches(1.2), Inches(4), Inches(0.4))
    heading_frame = heading_box.text_frame
    p = heading_frame.paragraphs[0]
    p.text = "五大超能力①"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLORS['primary_blue']
    
    box1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 
                                  Inches(5.5), Inches(1.8), 
                                  Inches(4), Inches(1.6))
    box1.fill.solid()
    box1.fill.fore_color.rgb = RGBColor(230, 245, 255)
    box1.line.color.rgb = COLORS['sky_blue']
    
    emoji1_box = slide.shapes.add_textbox(Inches(5.5), Inches(1.85), Inches(4), Inches(0.4))
    emoji1_frame = emoji1_box.text_frame
    p = emoji1_frame.paragraphs[0]
    p.text = "🧊 超冷晶片"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLORS['sky_blue']
    
    desc1_box = slide.shapes.add_textbox(Inches(5.7), Inches(2.3), Inches(3.6), Inches(1))
    desc1_frame = desc1_box.text_frame
    desc1_frame.word_wrap = True
    p = desc1_frame.paragraphs[0]
    p.text = "當食物太燙時自動降溫\n溫度≥40°C時啟動，≤37°C時停止"
    p.font.size = Pt(10)
    p.font.color.rgb = COLORS['dark_gray']
    
    box2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 
                                  Inches(5.5), Inches(3.6), 
                                  Inches(4), Inches(1.6))
    box2.fill.solid()
    box2.fill.fore_color.rgb = RGBColor(230, 245, 255)
    box2.line.color.rgb = COLORS['sky_blue']
    
    emoji2_box = slide.shapes.add_textbox(Inches(5.5), Inches(3.65), Inches(4), Inches(0.4))
    emoji2_frame = emoji2_box.text_frame
    p = emoji2_frame.paragraphs[0]
    p.text = "👁️ 紅外線眼睛"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLORS['sky_blue']
    
    desc2_box = slide.shapes.add_textbox(Inches(5.7), Inches(4.1), Inches(3.6), Inches(1))
    desc2_frame = desc2_box.text_frame
    desc2_frame.word_wrap = True
    p = desc2_frame.paragraphs[0]
    p.text = "看穿食物有多熱（不用戳）\n非接觸式即時量測食物內部溫度"
    p.font.size = Pt(10)
    p.font.color.rgb = COLORS['dark_gray']
    
    illus_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 
                                       Inches(0.5), Inches(1.2), 
                                       Inches(4.5), Inches(3.8))
    illus_box.fill.solid()
    illus_box.fill.fore_color.rgb = COLORS['primary_blue']
    illus_box.line.color.rgb = COLORS['sky_blue']
    
    illus_text = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(4.5), Inches(1))
    illus_frame = illus_text.text_frame
    p = illus_frame.paragraphs[0]
    p.text = "[致冷晶片 + 溫度下降動畫]"
    p.font.size = Pt(12)
    p.font.color.rgb = COLORS['white']
    p.alignment = PP_ALIGN.CENTER

def add_tech_slide_2(prs):
    """Slide 5: Core Tech - Alerts & LED"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 250, 240)
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = "核心技術 (2/2) — 警示系統"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLORS['primary_blue']
    
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.75), Inches(9), Inches(0.3))
    subtitle_frame = subtitle_box.text_frame
    p = subtitle_frame.paragraphs[0]
    p.text = "五大超能力②"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLORS['warning_yellow']
    
    statuses = [
        ("🔴", "紅燈警告", "暫停！太燙了！", "≥40°C", COLORS['danger_red']),
        ("🟡", "黃燈小心", "等一下...", "39–39.5°C", COLORS['warning_yellow']),
        ("🟢", "綠燈安全", "開吃！", "≤37°C", COLORS['safe_green'])
    ]
    
    x_positions = [0.6, 3.6, 6.6]
    for i, (emoji, title, desc, temp, color) in enumerate(statuses):
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 
                                     Inches(x_positions[i]), Inches(1.3), 
                                     Inches(2.6), Inches(3.8))
        box.fill.solid()
        box.fill.fore_color.rgb = color
        box.line.width = Pt(0)
        
        emoji_box = slide.shapes.add_textbox(Inches(x_positions[i]), Inches(1.5), 
                                             Inches(2.6), Inches(0.4))
        emoji_frame = emoji_box.text_frame
        p = emoji_frame.paragraphs[0]
        p.text = emoji
        p.font.size = Pt(36)
        p.alignment = PP_ALIGN.CENTER
        
        title_box = slide.shapes.add_textbox(Inches(x_positions[i]), Inches(2), 
                                             Inches(2.6), Inches(0.4))
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        p = title_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER
        
        temp_box = slide.shapes.add_textbox(Inches(x_positions[i]), Inches(2.5), 
                                            Inches(2.6), Inches(0.5))
        temp_frame = temp_box.text_frame
        p = temp_frame.paragraphs[0]
        p.text = temp
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER
        
        desc_box = slide.shapes.add_textbox(Inches(x_positions[i] + 0.1), Inches(3.2), 
                                            Inches(2.4), Inches(1.2))
        desc_frame = desc_box.text_frame
        desc_frame.word_wrap = True
        p = desc_frame.paragraphs[0]
        p.text = desc
        p.font.size = Pt(11)
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER
        
        alert_box = slide.shapes.add_textbox(Inches(x_positions[i]), Inches(4.6), 
                                             Inches(2.6), Inches(0.3))
        alert_frame = alert_box.text_frame
        p = alert_frame.paragraphs[0]
        p.text = "🔊 蜂鳴器" if i < 2 else "✅ 安全"
        p.font.size = Pt(10)
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER

def add_led_details_slide(prs):
    """Slide 6: LED System Details"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLORS['white']
    
    title_box = slide.shapes.add_textbox(Inches(0.3), Inches(0.15), Inches(9.4), Inches(0.5))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = "警示系統詳解 — 三色燈號故事"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = COLORS['primary_blue']
    
    panels = [
        {
            'title': '危險 (DANGER)',
            'emoji': '🔴',
            'color': COLORS['danger_red'],
            'temp': '≥40°C',
            'desc': '食物溫度過高，存在嚴重燙傷風險。致冷晶片持續降溫中，請勿進食。',
            'voice': '「警告溫度過高」'
        },
        {
            'title': '小心 (CAUTION)',
            'emoji': '🟡',
            'color': COLORS['warning_yellow'],
            'temp': '39–39.5°C',
            'desc': '溫度接近安全範圍，仍請繼續等待，暫勿讓幼童進食。',
            'voice': '「溫度升高，請小心」'
        },
        {
            'title': '安全 (SAFE)',
            'emoji': '🟢',
            'color': COLORS['safe_green'],
            'temp': '≤37°C',
            'desc': '食物已達安全入口溫度，蜂鳴器自動關閉，可以放心進食。',
            'voice': '「開吃」'
        }
    ]
    
    x_positions = [0.25, 3.45, 6.65]
    for idx, (panel, x_pos) in enumerate(zip(panels, x_positions)):
        panel_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 
                                          Inches(x_pos), Inches(0.85), 
                                          Inches(3.1), Inches(4.5))
        panel_box.fill.solid()
        panel_box.fill.fore_color.rgb = panel['color']
        panel_box.line.width = Pt(0)
        
        title_box = slide.shapes.add_textbox(Inches(x_pos), Inches(1), 
                                             Inches(3.1), Inches(0.4))
        title_frame = title_box.text_frame
        p = title_frame.paragraphs[0]
        p.text = f"{panel['emoji']} {panel['title']}"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER
        
        temp_box = slide.shapes.add_textbox(Inches(x_pos), Inches(1.5), 
                                            Inches(3.1), Inches(0.5))
        temp_frame = temp_box.text_frame
        p = temp_frame.paragraphs[0]
        p.text = panel['temp']
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER
        
        desc_box = slide.shapes.add_textbox(Inches(x_pos + 0.15), Inches(2.2), 
                                            Inches(2.8), Inches(1.5))
        desc_frame = desc_box.text_frame
        desc_frame.word_wrap = True
        p = desc_frame.paragraphs[0]
        p.text = panel['desc']
        p.font.size = Pt(9)
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER
        
        beeper_box = slide.shapes.add_textbox(Inches(x_pos), Inches(3.9), 
                                              Inches(3.1), Inches(0.3))
        beeper_frame = beeper_box.text_frame
        p = beeper_frame.paragraphs[0]
        p.text = "🔊 蜂鳴器" if idx < 2 else "✅ 關閉"
        p.font.size = Pt(9)
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER
        
        voice_box = slide.shapes.add_textbox(Inches(x_pos + 0.1), Inches(4.35), 
                                             Inches(2.9), Inches(0.8))
        voice_frame = voice_box.text_frame
        voice_frame.word_wrap = True
        p = voice_frame.paragraphs[0]
        p.text = panel['voice']
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER

def add_use_cases_slide(prs):
    """Slide 7: Use Cases"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(232, 245, 233)
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = "應用場景 — 誰需要暖口安心？"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLORS['primary_blue']
    
    scenarios = [
        ("🏠", "家裡", "家庭用餐安全保護", RGBColor(227, 242, 253)),
        ("🏫", "托嬰所", "照護者守護孩子健康", RGBColor(243, 229, 245)),
        ("🏥", "醫院", "長者照護與恢復支持", RGBColor(255, 243, 224)),
        ("🌍", "世界", "所有孩子都應該安全", RGBColor(252, 228, 236))
    ]
    
    positions = [(0.5, 1.2), (5.2, 1.2), (0.5, 3.4), (5.2, 3.4)]
    for (emoji, title, desc, bg_color), (x, y) in zip(scenarios, positions):
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 
                                     Inches(x), Inches(y), 
                                     Inches(4.2), Inches(1.8))
        box.fill.solid()
        box.fill.fore_color.rgb = bg_color
        box.line.color.rgb = COLORS['sky_blue']
        box.line.width = Pt(1)
        
        emoji_box = slide.shapes.add_textbox(Inches(x), Inches(y + 0.1), 
                                             Inches(4.2), Inches(0.4))
        emoji_frame = emoji_box.text_frame
        p = emoji_frame.paragraphs[0]
        p.text = emoji
        p.font.size = Pt(32)
        p.alignment = PP_ALIGN.CENTER
        
        title_box = slide.shapes.add_textbox(Inches(x), Inches(y + 0.5), 
                                             Inches(4.2), Inches(0.35))
        title_frame = title_box.text_frame
        p = title_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = COLORS['primary_blue']
        p.alignment = PP_ALIGN.CENTER
        
        desc_box = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.9), 
                                            Inches(3.8), Inches(0.7))
        desc_frame = desc_box.text_frame
        desc_frame.word_wrap = True
        p = desc_frame.paragraphs[0]
        p.text = desc
        p.font.size = Pt(11)
        p.font.color.rgb = COLORS['dark_gray']
        p.alignment = PP_ALIGN.CENTER

def add_sdgs_slide(prs):
    """Slide 8: Sustainable Development Goals"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLORS['primary_blue']
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.5))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = "永續發展 — 我們在乎的事"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLORS['white']
    
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.85), Inches(9), Inches(0.3))
    subtitle_frame = subtitle_box.text_frame
    p = subtitle_frame.paragraphs[0]
    p.text = "UN Sustainable Development Goals (SDGs)"
    p.font.size = Pt(12)
    p.font.color.rgb = RGBColor(200, 200, 200)
    
    sdgs = [
        ("3", "🏥", "健康與福祉", "保護孩子不被燙傷，降低食道癌風險"),
        ("9", "⚡", "科技創新", "用聰明的設計解決全球飲食安全問題"),
        ("12", "♻️", "永續消費", "減少醫療資源浪費，降低社會成本")
    ]
    
    x_positions = [0.8, 3.6, 6.4]
    for (num, emoji, title_zh, desc), x_pos in zip(sdgs, x_positions):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 
                                      Inches(x_pos), Inches(1.4), 
                                      Inches(2.6), Inches(3.7))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(30, 60, 180)
        card.line.width = Pt(0)
        
        num_box = slide.shapes.add_textbox(Inches(x_pos), Inches(1.6), 
                                           Inches(2.6), Inches(0.5))
        num_frame = num_box.text_frame
        p = num_frame.paragraphs[0]
        p.text = num
        p.font.size = Pt(48)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER
        
        emoji_box = slide.shapes.add_textbox(Inches(x_pos), Inches(2.2), 
                                             Inches(2.6), Inches(0.4))
        emoji_frame = emoji_box.text_frame
        p = emoji_frame.paragraphs[0]
        p.text = emoji
        p.font.size = Pt(28)
        p.alignment = PP_ALIGN.CENTER
        
        title_box = slide.shapes.add_textbox(Inches(x_pos + 0.1), Inches(2.75), 
                                             Inches(2.4), Inches(0.5))
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        p = title_frame.paragraphs[0]
        p.text = title_zh
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER
        
        desc_box = slide.shapes.add_textbox(Inches(x_pos + 0.15), Inches(3.65), 
                                            Inches(2.3), Inches(1.3))
        desc_frame = desc_box.text_frame
        desc_frame.word_wrap = True
        p = desc_frame.paragraphs[0]
        p.text = desc
        p.font.size = Pt(9)
        p.font.color.rgb = RGBColor(200, 220, 255)
        p.alignment = PP_ALIGN.CENTER

def add_team_slide(prs):
    """Slide 9: Team Introduction"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(243, 229, 245)
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = "我們的團隊 — 4位超新星"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLORS['primary_blue']
    
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.75), Inches(9), Inches(0.4))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.word_wrap = True
    p = subtitle_frame.paragraphs[0]
    p.text = "四位高中生，因為一個生活中的痛點，決定走出舒適圈，用科技守護每一口飲食的安全。"
    p.font.size = Pt(11)
    p.font.color.rgb = COLORS['dark_gray']
    
    members = [
        ("王婕羽", "想像力達人", COLORS['primary_blue']),
        ("王紫瑄", "技術專家", COLORS['sky_blue']),
        ("黃懷葦", "設計鬼才", COLORS['warning_yellow']),
        ("廖笙宏", "推廣大使", COLORS['safe_green'])
    ]
    
    x_positions = [0.8, 2.8, 4.8, 6.8]
    for (name, badge, color), x_pos in zip(members, x_positions):
        avatar = slide.shapes.add_shape(MSO_SHAPE.OVAL, 
                                       Inches(x_pos + 0.25), Inches(1.5), 
                                       Inches(1.3), Inches(1.3))
        avatar.fill.solid()
        avatar.fill.fore_color.rgb = color
        avatar.line.width = Pt(0)
        
        avatar_text = slide.shapes.add_textbox(Inches(x_pos + 0.25), Inches(1.6), 
                                              Inches(1.3), Inches(1.1))
        avatar_frame = avatar_text.text_frame
        avatar_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = avatar_frame.paragraphs[0]
        p.text = name[0]
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']
        p.alignment = PP_ALIGN.CENTER
        
        name_box = slide.shapes.add_textbox(Inches(x_pos), Inches(3), 
                                           Inches(1.8), Inches(0.4))
        name_frame = name_box.text_frame
        name_frame.word_wrap = True
        p = name_frame.paragraphs[0]
        p.text = name
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = COLORS['dark_bg']
        p.alignment = PP_ALIGN.CENTER
        
        badge_box = slide.shapes.add_textbox(Inches(x_pos), Inches(3.5), 
                                            Inches(1.8), Inches(0.5))
        badge_frame = badge_box.text_frame
        badge_frame.word_wrap = True
        p = badge_frame.paragraphs[0]
        p.text = badge
        p.font.size = Pt(9)
        p.font.color.rgb = color
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER

def add_closing_slide(prs):
    """Slide 10: Closing & CTA"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLORS['primary_blue']
    
    main_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.8), Inches(9), Inches(0.8))
    main_frame = main_box.text_frame
    main_frame.word_wrap = True
    p = main_frame.paragraphs[0]
    p.text = "暖口安心"
    p.font.size = Pt(52)
    p.font.bold = True
    p.font.color.rgb = COLORS['white']
    p.alignment = PP_ALIGN.CENTER
    
    tag_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.7), Inches(9), Inches(0.5))
    tag_frame = tag_box.text_frame
    p = tag_frame.paragraphs[0]
    p.text = "Keep Every Bite Safe!"
    p.font.size = Pt(28)
    p.font.color.rgb = COLORS['sky_blue']
    p.alignment = PP_ALIGN.CENTER
    
    cta_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(0.4))
    cta_frame = cta_box.text_frame
    p = cta_frame.paragraphs[0]
    p.text = "讓每一口都安全！ Make Every Bite Count!"
    p.font.size = Pt(18)
    p.font.color.rgb = COLORS['white']
    p.alignment = PP_ALIGN.CENTER
    
    buttons = [
        ("📱", "想了解更多？", "Learn More"),
        ("🎁", "想試試看？", "Try It Now"),
        ("🌟", "想加入我們？", "Join Us")
    ]
    
    y_pos = 3.2
    for emoji, text_zh, text_en in buttons:
        btn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 
                                     Inches(1.5), Inches(y_pos), 
                                     Inches(7), Inches(0.45))
        btn.fill.solid()
        btn.fill.fore_color.rgb = RGBColor(255, 255, 255)
        btn.line.width = Pt(1)
        btn.line.color.rgb = COLORS['white']
        
        btn_text = slide.shapes.add_textbox(Inches(1.5), Inches(y_pos + 0.05), 
                                           Inches(7), Inches(0.35))
        btn_frame = btn_text.text_frame
        btn_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = btn_frame.paragraphs[0]
        p.text = f"{emoji}  {text_zh} / {text_en}"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = COLORS['primary_blue']
        p.alignment = PP_ALIGN.CENTER
        
        y_pos += 0.55
    
    contact_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.8), Inches(9), Inches(0.6))
    contact_frame = contact_box.text_frame
    contact_frame.word_wrap = True
    p = contact_frame.paragraphs[0]
    p.text = "📧 csps.104195@chc.edu.tw  |  🌐 smart-thermal-bite.base44.app"
    p.font.size = Pt(11)
    p.font.color.rgb = COLORS['white']
    p.alignment = PP_ALIGN.CENTER
    
    footer_box = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(9), Inches(0.3))
    footer_frame = footer_box.text_frame
    p = footer_frame.paragraphs[0]
    p.text = "NOVA LAB 2026 | 感謝您的聆聽！ Thank You!"
    p.font.size = Pt(11)
    p.font.color.rgb = RGBColor(200, 200, 200)
    p.alignment = PP_ALIGN.CENTER

def main():
    """Main function"""
    print("🚀 開始生成暖口安心 PPT...")
    print("=" * 50)
    
    prs = create_presentation()
    
    print("📝 正在製作投影片...")
    print("  Slide 1: 標題頁")
    add_title_slide(prs)
    print("  Slide 2: 問題導入")
    add_problem_slide(prs)
    print("  Slide 3: 解決方案")
    add_solution_slide(prs)
    print("  Slide 4: 核心技術 (1/2)")
    add_tech_slide_1(prs)
    print("  Slide 5: 核心技術 (2/2)")
    add_tech_slide_2(prs)
    print("  Slide 6: 警示系統詳解")
    add_led_details_slide(prs)
    print("  Slide 7: 應用場景")
    add_use_cases_slide(prs)
    print("  Slide 8: 永續發展")
    add_sdgs_slide(prs)
    print("  Slide 9: 團隊介紹")
    add_team_slide(prs)
    print("  Slide 10: 結尾頁")
    add_closing_slide(prs)
    
    output_filename = "nova_lab_kids_presentation.pptx"
    prs.save(output_filename)
    
    print("=" * 50)
    print(f"✅ PPT 生成成功！")
    print(f"📁 檔案名稱: {output_filename}")
    print(f"📏 投影片數量: {len(prs.slides)} 張")
    print("=" * 50)
    print("\n💡 下一步:")
    print("  用 PowerPoint / Google Slides / Keynote 打開此檔案")
    print("  可自行調整內容、添加圖片、設定動畫")
    print("\n🎉 祝您發表會圓滿成功！")

if __name__ == "__main__":
    main()
