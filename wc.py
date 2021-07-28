#WORD COUNTER
#Ferrer, Mariñas, Tiu
#3ITD Samsantech

from collections import Counter
import numpy as np
import matplotlib.pyplot as plt
import string
import sys
import docx
import re


def main():
    print('Welcome to Samsantech Word Counter.\n',
          '\nWhat kind of word file are you going to use?\n',
          'A. .txt file\n',
          'B. .docx file\n')

    try:
        format = (input('Enter a letter only (A/B): '))
        
        if format.isalpha():
            #Transform upper to lower case
            format = format.lower()


            #.txt file
            if (format == 'a'):
                file_title = input('Enter the title of the document: ')
                file = input('Input the path of the word file: ')
                file = open(file,'r')

                #words + word count
                wc = {}

                for line in file:
                    line = line.strip()     #removes leading & trailing chars
                    line = line.translate(line.maketrans('', '', string.punctuation))
                    line = line.lower()
                    
                    rawWords = line.split()

                    stopwords = ['the', 'of', 'and', 'is','to','in','a','from','by',
                                 'that', 'with', 'this', 'as', 'an', 'are','its', 'at', 'for']

                    words = [ word for word in rawWords if word not in stopwords ]
                    
                    for word in words:
                        if word in wc:
                            wc[word] = wc[word] + 1
                        else:
                            wc[word] = 1

                #PRINT FORMAT
                print('\n{:^14}{:^14}'.format('Words','Count')) # format columns with headers

                #print every word in the wc
                for word in wc:
                    print('{:^14}{:^14d}'.format(word,wc[word]))     # format data and print

                print("")


                ## MATPLOTLIB PART
    
                counts = Counter(wc)
                #counts = dict(Counter(wc).most_common(10))      # first n entries

                labels, values = zip(*counts.items())

                # sort your values in descending order
                indSort = np.argsort(values)[::-1]

                # rearrange your data
                labels = np.array(labels)[indSort]
                values = np.array(values)[indSort]

                indexes = np.arange(len(labels))

                plt.bar(indexes, values, color='yellow')

                # add labels
                plt.xticks(indexes, labels, fontsize=4)

                plt.xlabel('Words')
                plt.ylabel('Frequency')
                plt.title(file_title)
                
                plt.show()
            
            #.docx file
            elif (format == 'b'):
                file_title = input('Enter the title of the document: ')
                file = input('Input the path of the word file: ')
                doc = docx.Document(file)

                #words + word count
                wc = {}
                text = []
                
                for para in doc.paragraphs:     #access each parag. in the doc
                    text.append(para.text)

                rawWords = re.findall(r'\w+', '\n'.join(text).lower())

                stopwords = ['the', 'of', 'and', 'is','to','in','a','from','by',
                                 'that', 'with', 'this', 'as', 'an', 'are','its', 'at', 'for']

                words = [ word for word in rawWords if word not in stopwords ]
                
                for word in words:
                    if word in wc:
                        wc[word] = wc[word] + 1
                    else:
                        wc[word] = 1
                
                print('\n{:^14}{:^14}'.format('Words','Count')) # format columns with headers

                #print every word in the wc
                for word in wc:
                    print('{:^14}{:^14d}'.format(word,wc[word]))     # format data and print

                print("")

                
                ## MATPLOTLIB PART
    
                counts = Counter(wc)
                #counts = dict(Counter(wc).most_common(10))      # first n entries

                labels, values = zip(*counts.items())

                # sort your values in descending order
                indSort = np.argsort(values)[::-1]

                # rearrange your data
                labels = np.array(labels)[indSort]
                values = np.array(values)[indSort]

                indexes = np.arange(len(labels))

                plt.bar(indexes, values, color='yellow')

                # add labels
                plt.xticks(indexes, labels, fontsize=3)

                plt.xlabel('Words')
                plt.ylabel('Frequency')
                plt.title(file_title)
                
                plt.show()
   
            else:
                print('Enter the correct input, A or B only.')

        else:
            print('Enter the correct input, A or B only.')
                
    except Exception as e:
        print(e)

main()


while True:
    try_again = input('\nTry again? Y/N: ')
    try_again = try_again.lower()
    
    if try_again == 'y':
        main()
    elif try_again == 'n':
        print('\nThank you for trying Samsantech Word Counter.')
        print('Ferrer, Mariñas, Tiu | 3ITD')
        sys.exit()
    else:
        try_again
