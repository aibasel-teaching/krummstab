#! /bin/sh
[ $# -ne 1 ] && echo "Provide the .xlsx file exported from ADAM as input." && exit 1
filebasename="$(basename -- "$1" '.xlsx')"

# Convert .xlsx file to .csv.
libreoffice --headless --convert-to csv $1 --outdir .
csvfile="${filebasename}.csv"

# For me, the csv file is encoded in iso-8859-1 while my shell uses an UTF-8
# encoding. This caused sed's '.' to not match special characters such as
# umlauts. Setting the LANG environment variable accordingly should fix this.
LANG=$(file -i ${csvfile} | sed 's/^.*=//')

# Remove first line.
sed -i '1d' ${csvfile}

# Parse first name, last name, and email and print in JSON entry format.
sed -i 's/\"\(..*\), \(..*\) \[\(.*\)\].*/\[\"\2\", \"\1\", \"\3\"\],/' ${csvfile}

# Sort.
sort -o ${csvfile} -- ${csvfile}

# Convert to UTF-8 and change extension to '.json'.
iconv -f ${LANG} -t UTF-8 < ${csvfile} > ${filebasename}.json

# Remove intermmediate .csv file.
rm ${csvfile}
