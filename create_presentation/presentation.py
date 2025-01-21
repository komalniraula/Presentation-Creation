import nltk
from nltk.corpus import stopwords 
import pdfplumber
from keybert import KeyBERT
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import requests
from bs4 import BeautifulSoup
import os.path
from pptx import Presentation
from create_presentation.powerpoint import create_powerpoint
from personal.settings import BASE_DIR

class content_presentation:

    def __init__(self, user_name, file_name, file_path) -> None:
        self.user_name = user_name
        self.file_name = file_name
        self.file_name = 'User_files/' + self.user_name + '/uploads/'+ self.file_name
        self.file_path = file_path
        self.used_images = []
        self.kw_model = KeyBERT()
        self.model = SentenceTransformer('paraphrase-MiniLM-L6-v2')        

    def keyword_extraction(self, full_text):    
        sentences = nltk.sent_tokenize(full_text)
        #print(sentences)
        freepik_search = []
        google_search = []
        for sentence in sentences:
            #sentence = sentence.lower()
            nouns = []
            proper_noun_main = []
            proper_noun = []
            #print(sentence)
            word_list = sentence.split()
            number_of_words = len(word_list)
            number_of_images = int(number_of_words / 3)
            print(sentence)
            key_word = self.kw_model.extract_keywords(sentence, keyphrase_ngram_range=(1, 2), stop_words='english', top_n=number_of_images)
            #print(key_word)
            #print('-------------------')
            words = nltk.word_tokenize(sentence)
            words = [word for word in words if word not in set(stopwords.words('english'))]            
            tagged = nltk.pos_tag(words)
            for (word, tag) in tagged:
                if tag == 'NN':
                    if tag != 'NNP' or tag != 'NNPS':
                        nouns.append(word)
                elif tag == 'NNP' or tag == 'NNPS':
                    proper_noun_main.append(word)

            proper_noun = [x.lower() for x in proper_noun_main]
            
            for w in key_word:
                wr, cos = w
                sp = wr.split()
                for k in sp:
                    if k in nouns:
                        if len(sp) == 2:
                            first_word = sp[0]
                            second_word = sp[1]
                            if first_word in nouns:
                                if second_word in nouns:
                                    if wr not in freepik_search:
                                        freepik_search.append(wr)
                                elif second_word not in nouns:
                                    if first_word not in freepik_search:
                                        freepik_search.append(first_word)
                            elif second_word in nouns:
                                if second_word not in freepik_search:
                                    freepik_search.append(second_word)

                        elif wr not in freepik_search:
                            freepik_search.append(wr)

                    elif k in proper_noun:
                        if len(sp) == 2:
                            first_word = sp[0]
                            second_word = sp[1]
                            if first_word in proper_noun:
                                if second_word in proper_noun:
                                    if wr not in google_search:
                                        first_word_cap = first_word.capitalize()
                                        second_word_cap = second_word.capitalize()
                                        if first_word_cap in proper_noun_main and second_word_cap in proper_noun_main:
                                            google_search.append(first_word)
                                            google_search.append(second_word)

                                        elif first_word_cap in proper_noun_main and second_word_cap not in proper_noun_main:
                                            google_search.append(wr)

                                elif second_word not in proper_noun:
                                    if first_word not in google_search:
                                        google_search.append(first_word)

                            elif second_word in proper_noun:
                                if second_word not in google_search:
                                    google_search.append(second_word)

                        elif wr not in google_search:
                            google_search.append(wr)

        google_search = list(set(google_search))
        return freepik_search, google_search

    def title_text(self, page):
        t = page.extract_text()
        return t

    def get_text_for_ppt(self, page):
        clean_text = page.filter(lambda obj: (obj["object_type"] == "char" and "Bold" in obj["fontname"]))
        text = clean_text.extract_text()
        return text

    def sentences_for_ppt(self, text):
        sentences_for_presentation = []
        i = 1
        sentences = nltk.sent_tokenize(text)
        for sentence in sentences:
            sentence = sentence.replace('\n',"")
            sentences_for_presentation.append(sentence)
            i += 1
        return sentences_for_presentation
    
    def get_freepik_imglist(self, keyword):
        results = {}
        title_list = []
        for i in range(4):
            page = i+1
            url = 'https://www.freepik.com/search?format=search&page={pag}&query={key}&type=vector'.format(pag = page, key = keyword)
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'}
            page = requests.get(url, headers=headers)
            soup = BeautifulSoup(page.text, 'html.parser')
            for div in soup.find_all('p', {'class':"cleaned-filters"}):
                return results, title_list 
                
            for div in soup.find_all('a', {'class':"showcase__link"}):
                img = div.find('img', alt=True)
                title = img['alt']
                title = title.lower()
                link = div['href']
                
                every_word = link.split('.')

                if 'freepik' in every_word:
                    
                    lin = link.split("vector/",1)[1]
                    final_til = title + ' ' + lin[:-4].replace('-', ' ').lower()
                    
                    if final_til not in title_list:
                        title_list.append(final_til)
                    
                    results[final_til] = link
        return results, title_list 

    def embeddings(self, title_list, full_text):
        embeddings = self.model.encode(title_list)
        embeddings_for_main = self.model.encode(full_text)
        embeddings_list = {}
        for title, embedding in zip(title_list, embeddings):
            embeddings_list[title] = embedding
        return embeddings_for_main, embeddings_list
    
    def get_cosine(self, embeddings_for_main, embeddings_list):
        cosine_value = {}
        for key in embeddings_list:
            c = cosine_similarity([embeddings_for_main], [embeddings_list[key]])
            k = float(c[0])
            cosine_value[key] = k
        cosine_value = {k: v for k, v in sorted(cosine_value.items(), key=lambda item: item[1], reverse = True)}
        return cosine_value
    
    def get_image_link(self, cosine_value, used_images, results):
        image_link = 0
        for key in cosine_value:
            if key not in self.used_images:
                image_link = results[key]
                used_images.append(key)
                return image_link, used_images
        return image_link, used_images
    
    def get_image(self, url):
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'}
        page = requests.get(url, headers=headers)
        soup = BeautifulSoup(page.text, 'html.parser')
        for x in soup.findAll('div', {'class':"detail__gallery detail__gallery--vector alignc"}):
            img = x.find('img', alt=True)
            source = img['src']
        return source
    
    def save_image(self, img, file_path): 
        img_name = img.split('.jpg')[0]
        img_name = img_name.split('vector/')[1]
        img_name = img_name[:155]+'.jpg'
        completeName = file_path + '/' + img_name
        full_file_location = os.path.join(BASE_DIR, completeName)
        #full_file_path = Path(completeName)
        print('complete name is: ', full_file_location)
        print('img url is : ', img)

        headers={'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Cafari/537.36'}
        pic = requests.get(img, headers=headers)
        with open(full_file_location, 'wb+') as photo:
            photo.write(pic.content)
        return full_file_location

    def img_freepik(self, freepik_search, full_text, used_images, file_path):
        complete_image_path = {}
        for image_title in freepik_search:
            try:
                results, title_list = self.get_freepik_imglist(image_title)
                embeddings_for_main, embedding_list = self.embeddings(title_list, full_text)
                cosine_value = self.get_cosine(embeddings_for_main, embedding_list)
                image_link, used_images = self.get_image_link(cosine_value, used_images, results)
                if image_link != 0:
                    image_source = self.get_image(image_link)
                    complete_image_loc = self.save_image(image_source, file_path)
                    complete_image_path[complete_image_loc] = image_link
            except:
                pass
        return(complete_image_path, used_images)

    def texts(self):
        prs = Presentation()
        with pdfplumber.open(self.file_name) as pdf:
            totalpages = len(pdf.pages)
            print(totalpages)
            for i in range(totalpages):
                page = pdf.pages[i]
                full_text = self.title_text(page)
                freepik_search, google_search = self.keyword_extraction(full_text)
                ppt_texts = self.get_text_for_ppt(page)
                sentences_to_keep = self.sentences_for_ppt(ppt_texts)
                images, self.used_images = self.img_freepik(freepik_search, full_text, self.used_images, self.file_path)
                powerpoint = create_powerpoint(prs, sentences_to_keep, images)
                prs = powerpoint.presentation()
        return prs


    

    

