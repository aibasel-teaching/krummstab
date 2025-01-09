# The "convert" command comes from ImageMagick. Starting with ImageMagick 7, the
# command will change to "magick".
# As of 2025-01-09 the convert commands don't work properly and refuse to render
# parts of the image that are present in the PDFs. Doing the conversion with a
# random online tool works though.
logo-pdf:
	pdflatex logo.tex

logo-png: logo-pdf
	convert -density 100 logo.pdf -quality 90 -background white -alpha remove -alpha off logo.png

logo-jpg: logo-pdf
	convert -density 100 logo.pdf -quality 100 logo.jpg

banner-pdf:
	pdflatex banner.tex
