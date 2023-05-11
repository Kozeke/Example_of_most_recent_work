It's an example of the code which creates documents using the library package PhpWord (print .docx extension documents in Laravel) and in some cases extends this library.
The front side had an editor, which creates and visualizes each line of the document based on settings.
Basically in the document there can be records, table and signatures. 
Records - consist of title and field, usually title which is some text and field which is some underscore line which will be filled later. 
For records and fields, there are many settings like: alignment, fontSize, fontFamily, printLine, size and etc.
Most interesting feature is when sizeType is "content", it means size of the cell that is for the title depends on it's content. For instance, If the title "Name, surname of the respondent" should be printed, in PhpWord library you have to give a size for the cell which will have the text. And in PhpWord there is no function to know the size of the word. So we have created a function for finding the size of the word.    
