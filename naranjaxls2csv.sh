function xls2csv {
	unoconv -f csv ${1}
	#unoconv -f csv -e FilterOptions=124,34,76 ${1}
	CSV=`echo $1 | sed 's/\.xls$/\.csv/'`
	grep "^[0-9][0-9]/[0-9][0-9]/20[0-9][0-9]," ${CSV} > ${CSV}.out.csv
}

if [ $# -eq 1 ]
then
	xls2csv $1
else
	ls *.xls | while read F
	do
		xls2csv $F
	done
fi
