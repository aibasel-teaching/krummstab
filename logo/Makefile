pdf:
	pdflatex logo.tex

# The convert command comes with ImageMagick.
png: pdf
	convert -density 100 logo.pdf -quality 90 logo.png

jpg: pdf
	convert -density 100 logo.pdf -quality 100 logo.jpg
