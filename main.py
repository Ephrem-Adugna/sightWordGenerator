from docx import Document
import re
from collections import Counter

def get_top_n_words(file_path, n):
    # Read the file with UTF-8 encoding
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()
    
    # Use regular expressions to extract words (considering Amharic characters)
    words = re.findall(r'\b[\u1200-\u137F]+\b', text)
    
    # Count the frequency of each word
    word_counts = Counter(words)
    
    # Get the top n most common words
    top_n_words = word_counts.most_common(n)
    return top_n_words

# Example usage
file_path = 'textInput.txt'  # Replace with your file path
n = 63  # Number of top words to retrieve
top_words = get_top_n_words(file_path, n)



document = Document()


i=0
row=0
wordInd=0
for r in range(7):
    table = document.add_table(rows=0, cols=3)
    table.style = 'TableGrid'
    row_cells = table.add_row().cells
    for word, count in top_words:
        if row ==2 and i ==3:
            row=0
            i=0
            document.add_page_break()
            break 
        if i == 3:
            row_cells = table.add_row().cells
            i=0
            row+=1
        row_cells[i].text = top_words[wordInd][0]
        wordInd+=1
        i+=1
        

document.save('sightWords.docx')