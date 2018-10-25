function xls2csv {
	unoconv -f csv ${1}
	CSV=`echo $1 | sed 's/\.xls$/\.csv/'`
	grep "^[0-9][0-9]/[0-9][0-9]/20[0-9][0-9]," ${CSV} > ${CSV}.out.csv
}

function xls2ofx {
	OUTPUT=${1}.out.ofx

	# XLS to CSV, output is in ${CSV}.out.csv
	xls2csv ${1}

	# Extract the position of the fields in the CSV file
	HEADERTAG="FECHA VALOR,"
	grep -B 4 "$HEADERTAG" ${CSV} > ${CSV}.header
	FECHA=`grep "$HEADERTAG" ${CSV}.header | awk -F, '{ for (i=1; i<=NF; i++) { if ($i == "FECHA VALOR") print i; } }'`
	DESCRIPCION=`grep "$HEADERTAG" ${CSV}.header | awk -F, '{ for (i=1; i<=NF; i++) { if (($i == "DESCRIPCION") || ($i == "DESCRIPCIÓN")) print i; } }'`
	IMPORTE=`grep "$HEADERTAG" ${CSV}.header | awk -F, '{ for (i=1; i<=NF; i++) { if ($i == "IMPORTE (€)") print i; } }'`

	# Extract the account Id
	ACCTID=`grep "Número de cuenta:," ${CSV}.header | awk -F, '{ print $4; }'`
    if [[ $ACCTID == "" ]]
    then
        ACCTID=`grep "Número de tarjeta:," ${CSV}.header | awk -F, '{ print $3; }'`
    fi
	BANKID="ING DIRECT"
	rm ${CSV}.header
	rm ${CSV}

	echo "FECHA       = "$FECHA
	echo "DESCRIPCION = "$DESCRIPCION
	echo "IMPORTE     = "$IMPORTE
	echo "ACCTID      = "$ACCTID

	NARANJA=`echo $1 | grep "CuentaNARANJA"`
    if [[ $NARANJA == "" ]]
    then
        ACCTTYPE="CHECKING"
    else
        ACCTTYPE="SAVINGS"
    fi
	echo "ACCTTYPE    = "$ACCTTYPE

	#exit

	# Print header (all lines in the template until "___STMTTRN___")
	IFS=''
	while read LINE
	do
		if [ "$LINE" == "          ___STMTTRN___" ] || [[ $FOUND -eq 1 ]]
		then
			FOUND=1
		else
			echo $LINE | sed "s/<ACCTID>ACCTID/<ACCTID>${ACCTID}/" | sed "s/<BANKID>BANKID/<BANKID>${BANKID}/"  | sed "s/<ACCTTYPE>ACCTTYPE/<ACCTTYPE>${ACCTTYPE}/"
		fi
	done < naranjaxlstemplate.ofx > $OUTPUT

	# Print body
	COUNTER=1
	#cat ${CSV}.out.csv | tail -n +1 | awk -vFPAT='[^,]*|"[^"]*"' '{
	cat ${CSV}.out.csv | tail -n +1 | awk -vFPAT='[^,]*|"[^"]*"' -v FECHA=$FECHA -v DESCF=$DESCRIPCION -v IMPORTE=$IMPORTE '{
			DATE = $FECHA
			sub(" Jan ", " 01 ", DATE);
			sub(" Feb ", " 02 ", DATE);
			sub(" Mar ", " 03 ", DATE);
			sub(" Apr ", " 04 ", DATE);
			sub(" May ", " 05 ", DATE);
			sub(" Jun ", " 06 ", DATE);
			sub(" Jul ", " 07 ", DATE);
			sub(" Aug ", " 08 ", DATE);
			sub(" Sep ", " 09 ", DATE);
			sub(" Oct ", " 10 ", DATE);
			sub(" Nov ", " 11 ", DATE);
			sub(" Dec ", " 12 ", DATE);

			if (substr(DATE, 2, 1) == " ")
				DATE = "0"DATE;

			DATE = substr(DATE, 7, 4)""substr(DATE, 4, 2)""substr(DATE, 1, 2)"000000.000";

			AMOUNT=$IMPORTE;
			sub("\"", "", AMOUNT);
			sub("\"", "", AMOUNT);

			DESC=$DESCF;
			sub("\"", "", DESC);
			sub("\"", "", DESC);
			sub("ó", "o", DESC);

			print DATE"Ẍ"DESC"Ẍ"AMOUNT;
		}' | while read T
	do
		DATE=`echo $T | awk -F"Ẍ" '{print $1}'`
		DESC=`echo $T | awk -F"Ẍ" '{print $2}'`
		AMOUNT=`echo $T | awk -F"Ẍ" '{print $3}' | sed 's/,//'`
		FITID=`date +%s`$COUNTER
		echo "          <STMTTRN>"
		echo "            <TRNTYPE>PAYMENT"
		echo "            <DTPOSTED>${DATE}[+1:CET]"
		echo "            <DTAVAIL>${DATE}[+1:CET]"
		echo "            <TRNAMT>$AMOUNT"
		echo "            <FITID>$FITID"
		echo "            <NAME>$DESC"
		echo "            <MEMO>$DESC"
		echo "          </STMTTRN>"
		((COUNTER=COUNTER+1))
	done >> $OUTPUT

	# Print footer (all lines in the template from "___STMTTRN___")
	IFS=''
	FOUND=0
	while read LINE
	do
		if [[ $FOUND -eq 1 ]]
		then
			echo $LINE
		fi

		if [ "$LINE" == "          ___STMTTRN___" ]
		then
			FOUND=1
		fi
	done < naranjaxlstemplate.ofx >> $OUTPUT
}

if [ $# -eq 1 ]
then
	#xls2csv $1
	xls2ofx $1 
else
	ls *.xls | while read F
	do
		#xls2csv $F
		xls2ofx $F
	done
fi
