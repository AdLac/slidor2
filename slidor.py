import os
import openai
import csv
import time
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
import streamlit as st

# OpenAI API Key
openai.api_key = os.getenv("OPENAI_API_KEY") or st.secrets["OPENAI_API_KEY"]

#Password
def check_password():
    """Returns `True` if the user had the correct password."""

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct.
        return True

if check_password():
    st.write("Here goes your normal Streamlit app...")
    st.button("Click me")


def main():
    # Titre
    st.title('Slide Generator :sunglasses:')

    # Load presentation template
    prs = Presentation('template nexus.pptx')
    # Define slide layout
    slide_layout = prs.slide_layouts[12]

    # Load keywords from file
    st.subheader('Générer les slides, 1 ligne = 1 titre de slide')
    contexte = st.text_area("Le contexte permets d'obtenir des résultats plus précis,'tu travailles pour ce client...'")
    keywords = st.text_area('1 ligne, 1 titre')

    if keywords:
        keywords = keywords.split("\n")
        counter = 0
        rows = []
        for keyword in keywords:
            keyword = keyword.strip()
            counter = counter + 1

            # Create prompt for OpenAI completion
            prompt2 = (
                    f"En te basant sur ce contexte: {contexte}  = Rédige le titre d'une slide ainsi que un commentaire complet et détaillé de slide powerpoint  à partir de cette idée:  \"" + keyword +"\", utilise le format suivant pour le titre <T>titre-généré-ici</T> et le format suivant pour le commentaire <C>commentaire-généré-ici</C>"
                      )

            # Get content from OpenAI
            response = openai.Completion.create(
                engine="text-davinci-003",
                prompt=prompt2,
                temperature=0.7,
                max_tokens=500,
                top_p=1,
                frequency_penalty=0,
                presence_penalty=0
            )

            content = response.choices[0].text
            title, body = content.split("<T>")[1].split("</T>")[0], content.split("<C>")[1].split("</C>")[0]

            rows.append([keyword, title, body])

            slide = prs.slides.add_slide(slide_layout)

            Presentation_Title = slide.placeholders[0]
            Presentation_Title.text = title

            #Add subtitle to slide
            Presentation_Subtitle = slide.placeholders[11]
            Presentation_Subtitle.text = title

            #Add body to slide
            Presentation_body = slide.placeholders[21]
            Presentation_body.text = body

            #Print progress
            print(str(counter) + " complete.")

            #Sleep to avoid spamming the API
            time.sleep(0.5)

        #Save the modified presentation
        st.text(str(counter) + " complete.")
        prs.save('template nexus modified.pptx')
        st.success("The modified presentation has been saved!")

        with open("template nexus modified.pptx", "rb") as file:
            btn = st.download_button(
                label="Download image",
                data=file,
                file_name="template nexus modified.pptx",
                mime='application/octet-stream',
            )
if __name__ == '__main__':
    main()