import collections.abc
import config
import requests

assert collections
import tkinter as tk
from tkinter import ttk

from pptx import Presentation
from pptx.util import Inches, Pt
import openai
from openai import OpenAI
from io import BytesIO

# API Token
openai.api_key = config.API_KEY
client = OpenAI(
    # defaults to os.environ.get("OPENAI_API_KEY")
    api_key=openai.api_key,
)


def slide_generator(text, prs):
    prompt = f"Summarize the following text to a DALL-E image generation " \
             f"prompt: \n {text}"

    dlp = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "user", "content": "I will ask you a question"},
            {"role": "assistant", "content": "Ok"},
            {"role": "user", "content": f"{prompt}"}
        ],
        max_tokens=250,
        n=1,
        stop=None,
        temperature=0.8
    )
    print(dir(dlp.choices[0]))

    dalle_prompt = dlp.choices[0].message.content.strip()

    response = client.images.generate(
        model="dall-e-3",
        prompt=dalle_prompt + " Style: digital art",
        quality="standard",
        n=1,
        size="1024x1024"
    )
    image_url = response.data[0].url

    prompt = f"Create a bullet point text for a Powerpoint" \
             f"slide from the following text: \n {text}"
    ppt = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "user", "content": "I will ask you a question"},
            {"role": "assistant", "content": "Ok"},
            {"role": "user", "content": f"{prompt}"}
        ],
        max_tokens=1024,
        n=1,
        stop=None,
        temperature=0.8
    )
    ppt_text = ppt.choices[0].message.content.strip()

    prompt = f"Create a title for a Powerpoint" \
             f"slide from the following text: \n {text}"
    ppt = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "user", "content": "I will ask you a question"},
            {"role": "assistant", "content": "Ok"},
            {"role": "user", "content": f"{prompt}"}
        ],
        max_tokens=1024,
        n=1,
        stop=None,
        temperature=0.8
    )
    ppt_header = ppt.choices[0].message.content.strip()

    # Add a new slide to the presentation
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    response = requests.get(image_url)
    img_bytes = BytesIO(response.content)
    slide.shapes.add_picture(img_bytes, Inches(1), Inches(1))

    # Add text box
    txBox = slide.shapes.add_textbox(Inches(3), Inches(1),
                                     Inches(4), Inches(1.5))
    tf = txBox.text_frame
    tf.text = ppt_text

    title_shape = slide.shapes.title
    title_shape.text = ppt_header


def update_progress_bar(current, total):
    progress['value'] = (current / total) * 100
    app.update_idletasks()  # Update the UI


def get_slides():
    text = text_field.get("1.0", "end-1c")
    paragraphs = text.split("\n\n")
    prs = Presentation()
    width = Pt(1920)
    height = Pt(1080)
    prs.slide_width = width
    prs.slide_height = height

    total_paragraphs = len(paragraphs)
    for i, paragraph in enumerate(paragraphs, start=1):
        slide_generator(paragraph, prs)
        update_progress_bar(i, total_paragraphs)

    prs.save("my_presentation.pptx")
    progress['value'] = 0  # Reset the progress bar after completion

    prs.save("my_presentation.pptx")


# Create the main window
app = tk.Tk()
app.title("Create PPT Slides")
app.geometry("800x600")

# Create the text field
text_field = tk.Text(app)
text_field.pack(fill="both", expand=True)
text_field.configure(wrap="word", font=("Arial", 12))
text_field.focus_set()

# Ensure the Progressbar is created after the main window `app` is defined
progress = ttk.Progressbar(app, orient="horizontal", length=300, mode="determinate")
progress.pack(pady=20)
# Create the button to create slides
create_button = tk.Button(app, text="Create Slides", command=get_slides)
create_button.pack()

app.mainloop()
