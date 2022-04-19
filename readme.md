<meta name='keywords' content='Austria, Österreich, Post AG, Brief, letter, template, reportlab, python'>

# AT Post Letter Template

This repository contains a template that mimics the Austrian [Post AG - Vorlage mit Absender](https://www.tages-post.at/tp/fe/assets/files/Tages-Post%20Vorlage-Musterbrief_Standard_mit_Absender.docx) letter template using ReportLab in Python. The main features are:

1. Create a .pdf letter in A4 respecting the [letter design settings](https://www.tages-post.at/tp/fe/templates) defined by the Post AG.
2. Iteration to create a .pdf with multiple pages (one letter per page), taking recipients from a given dataframe (an example is provided in the [at-post-letter-template.py](at-post-letter-template.py) code).
3. Fully personalizable: each text block is divided into frames; for each frame, it is possible to change the font and font size, add boundary to the text frames, change text colors, etc. Also specific words/sentences can be formatted (bold, italic, etc.) using simple HTML.

### Python dependencies

```python -m pip install pandas reportlab```

### Output

The [output.pdf](examples/output.pdf) gives the output of the [at-post-letter-template.py](at-post-letter-template.py) code.

<p align="center">
<img src="examples/outputa.png" alt="Output" width=510 high=720>
</p>

## Documentation

[Post AG Österreich - Briefgestaltung](https://www.tages-post.at/tp/fe/templates)
