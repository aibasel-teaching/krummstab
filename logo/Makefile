pdf:
	pdflatex logo.tex

# The convert command comes with ImageMagick.
png: pdf
	convert -density 100 logo.pdf -quality 90 -background white -alpha remove -alpha off logo.png

jpg: pdf
	convert -density 100 logo.pdf -quality 100 logo.jpg
