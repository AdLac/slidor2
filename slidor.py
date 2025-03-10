import openai
import time
import pandas as pd
from pptx import Presentation
import streamlit as st

# OpenAI API Key
client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

def main():
    # Titre
    st.title('Slide Generator 😎')

    # Load presentation template
    prs = Presentation('template nexus.pptx')
    slide_layout = prs.slide_layouts[12]

    # Load keywords from file
    st.subheader('Générer les slides, 1 ligne = 1 titre de slide')
    contexte = st.text_area("Le contexte permet d'obtenir des résultats plus précis, ex: 'tu travailles pour ce client...'")
    keywords = st.text_area('1 ligne, 1 titre')

    if keywords:
        keywords = keywords.split("\n")
        counter = 0
        rows = []

        for keyword in keywords:
            keyword = keyword.strip()
            if not keyword:
                continue
            
            counter += 1
            
            # Create prompt for OpenAI completion
            prompt2 = (
                f"{contexte} Rédige un titre ainsi qu'un commentaire complet et détaillé à partir de cette idée :  \"{keyword}\"."
                " Utilise le format suivant pour le titre <T>titre-généré-ici</T> et pour le commentaire <C>commentaire-généré-ici</C>."
            )

            # Get content from OpenAI
            response = client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": prompt2}
                ],
                temperature=0.7,
                max_tokens=500,
                top_p=1,
                frequency_penalty=0,
                presence_penalty=0
            )

            content = response.choices[0].message.content
            try:
                title = content.split("<T>")[1].split("</T>")[0]
                body = content.split("<C>")[1].split("</C>")[0]
            except IndexError:
                st.error("Erreur dans la réponse AI, format incorrect.")
                continue

            rows.append([keyword, title, body])

            slide = prs.slides.add_slide(slide_layout)
            slide.placeholders[0].text = title
            slide.placeholders[11].text = title
            slide.placeholders[21].text = body

            print(f"{counter} complet.")
            time.sleep(0.5)  # Pause pour éviter de spammer l'API

        st.text(f"{counter} complet.")
        prs.save('template_nexus_modified.pptx')
        st.success("La présentation modifiée a été enregistrée !")

        with open("template_nexus_modified.pptx", "rb") as file:
            st.download_button(
                label="Télécharger la présentation",
                data=file,
                file_name="template_nexus_modified.pptx",
                mime='application/octet-stream'
            )

if __name__ == '__main__':
    main()
