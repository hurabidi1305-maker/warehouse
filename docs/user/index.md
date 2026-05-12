from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

prs = Presentation()

def add_bullet_slide(prs, title_text, bullets):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = title_text
    tf = slide.placeholders[1].text_frame
    tf.clear()
    for bullet in bullets:
        p = tf.add_paragraph()
        p.text = bullet
    return slide

# Title Slide
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "Interest Rate Theory and Time Preference"
slide.placeholders[1].text = "Banking Systems Report\n[Your Names & Date]"

add_bullet_slide(prs, "Presentation Outline", [
    "1. Introduction to Time Preference",
    "2. Time Preference Theory of Interest",
    "3. Other Interest Rate Theories",
    "4. Determinants of Interest Rates",
    "5. Graphical Illustrations",
    "6. Relevance to Banking System",
    "7. Conclusion & Key Takeaways"
])

add_bullet_slide(prs, "1. Introduction to Time Preference", [
    "• General preference for present goods over future goods",
    "• Positive time preference is the norm",
    "• Explains the Time Value of Money",
    "• Subjective — varies by age, income, culture, expectations",
    "• Lower societal time preference → Higher saving & growth"
])

add_bullet_slide(prs, "2. Time Preference Theory of Interest", [
    "• Major contributor: Eugen von Böhm-Bawerk",
    "• Interest emerges from preference for present over future",
    "• Pure interest rate = societal rate of time preference",
    "• Natural compensation for waiting, not exploitation"
])

add_bullet_slide(prs, "3. Major Interest Rate Theories", [
    "• Loanable Funds Theory (Supply of savings vs Demand)",
    "• Productivity / Classical Theory",
    "• Keynesian Liquidity Preference",
    "• Modern synthesis: Time preference + Productivity + Risk"
])

add_bullet_slide(prs, "4. Key Determinants of Interest Rates", [
    "• Time Preference (core subjective factor)",
    "• Marginal Productivity of Capital",
    "• Inflation Expectations (Fisher Effect)",
    "• Risk, Liquidity, and Government Policy"
])

# Graph Slides (Add your images manually)
slide = prs.slides.add_slide(prs.slide_layouts[5])
title = slide.shapes.title
title.text = "5. Loanable Funds Market"
txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5))
tf = txBox.text_frame
tf.text = "Insert Loanable Funds Graph here\n\nSupply (Savers) ↑ with higher interest rates\nDemand (Borrowers) ↓ with higher rates\nEquilibrium rate balances saving & investment"

slide = prs.slides.add_slide(prs.slide_layouts[5])
title = slide.shapes.title
title.text = "Time Preference & Discounting"
txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5))
tf = txBox.text_frame
tf.text = "Insert Discounting Graph here\n\nShows how future value is discounted back to present"

add_bullet_slide(prs, "Time Preference Examples", [
    "High Time Preference (15–25%): Heavy borrowers",
    "Moderate (5–10%): Average consumers",
    "Low Time Preference (2–5%): Long-term savers & investors",
    "Historical real safe rate ≈ 2–4%"
])

add_bullet_slide(prs, "Relevance to Banking System", [
    "• Banks act as intermediaries between savers and borrowers",
    "• Profit from interest rate spread (Net Interest Margin)",
    "• Central banks influence short-term rates",
    "• Policy distortions can create bubbles and crises"
])

add_bullet_slide(prs, "Conclusion & Key Takeaways", [
    "• Interest rates fundamentally reflect human time preference",
    "• They coordinate saving, investment, and consumption over time",
    "• Crucial concept for understanding modern banking & monetary policy",
    "• Questions?"
])

prs.save('Interest_Rate_Theory_Time_Preference.pptx')
print("Presentation created successfully!")
