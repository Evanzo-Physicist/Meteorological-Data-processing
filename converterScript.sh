#!/bin/bash
### Please edit the "path", "file", and "sheetnumber" (begin counting the sheet number from 0,1,2,...)
path=/... #path
file='a'	#filename
let sheetnumber=3 #numbering sequence of excel sheets: 0,1,2,..
### Navigate to the directory set above
cd $path
### START COMPUTING AVERAGE PER MINUTE 
### convert the original excel file of format .xlsx that has your data into a .csv file
ssconvert -S ${file}.xlsx ${file}_sheet.csv
### Extract the relevant columns (Date, time, data) into a .dat (DATA file) 
### having in mind that .csv file has a "," separating the columns and delete the first 3 rows
awk 'BEGIN { FS=","; OFS"\t"; }{ print $1,$2,$23}' ${file}_sheet.csv.${sheetnumber} > ${file}_raw.dat
sed 1,3d ${file}_raw.dat > ${file}.dat
### Scan the values of the first row, specifically date DD,MN,YEAR, and time HH,MM,SS
let DD=$(sed -n 1p ${file}.dat | awk '{print substr($1,1,2)}' | bc)
let MN=$(sed -n 1p ${file}.dat | awk '{print substr($1,4,2)}' | bc)
let YEAR=$(sed -n 1p ${file}.dat | awk '{print substr($1,7,4)}' | bc)
let HH=$(sed -n 1p ${file}.dat | awk '{print substr($2,1,2)}' | bc)
let MM=$(sed -n 1p ${file}.dat | awk '{print substr($2,4,2)}' | bc)
let SS=$(sed -n 1p ${file}.dat | awk '{print substr($2,7,2)}' | bc)
### Compute the average of the first minute
### but first we determine the number of lines from the SS' second to 59th second.
let lines=60-${SS}
AvgPerMin=$(sed -n 1,${lines}p ${file}.dat | awk '{ total += $3; n++ } END { printf "%.2f",total/n}')
### Enter first the heading title to the Average-Per-Minute .dat file where TAB separates the columns
### Then append the first row having the Date, HH:MM, and AvgPerMin into the Average-Per-Minute .dat file
echo -e "Date\tTime\tAvg PT per min (W)" > ${file}_AvgPerMin.dat
echo -e "${DD}/${MN}/${YEAR}\t${HH}:${MM}\t${AvgPerMin}" >> ${file}_AvgPerMin.dat
### Having computed and entered the average of the first minute, first increment by 1 the "lines"
### variable. Then scan the values of DD,MN,YEAR,HH,MM
let lines=${lines}+1
let DD=$(sed -n ${lines}p ${file}.dat | awk '{print substr($1,1,2)}' | bc)
let MN=$(sed -n ${lines}p ${file}.dat | awk '{print substr($1,4,2)}' | bc)
let YEAR=$(sed -n ${lines}p ${file}.dat | awk '{print substr($1,7,4)}' | bc)
let HH=$(sed -n ${lines}p ${file}.dat | awk '{print substr($2,1,2)}' | bc)
let MM=$(sed -n ${lines}p ${file}.dat | awk '{print substr($2,4,2)}' | bc)
let SS=$(sed -n ${lines}p ${file}.dat | awk '{print substr($2,7,2)}' | bc)
### Count the number of records (total_lines) in ${file}.dat file for purposes of FOR-loop
wc -l ${file}.dat > temporaryfile.out
read total_lines filename < temporaryfile.out
rm -f temporaryfile.out
### FOR loop is initialized with the value of ${lines} and incremented by 60 in every loop
### Inside the loop, compute the average of a whole minute and append the results to .dat file
### Series of IF statements to increment, MM, HH, DD, MN, and YEAR. 
### MM(0 to 59), HH(0-23), DD(1 to some value depending on the month in question), MN(1 to 12), 
### YEAR(2021, next is 2022, - no explanation needed)
for ((i=${lines}; i<=${total_lines}; i+=60));
do
	### variable to determine the range of lines to do averaging. Call it upper limit, if you will.
	let lastsecond=${i}+59
	### Range in this case is from ${i} to ${lastsecond}
	AvgPerMin=$(sed -n ${i},${lastsecond}p ${file}.dat | awk '{ total += $3; n++ } END { printf "%.2f",total/n}')
	### Append Date, HH:MM, and AvgPerMin into the Average-Per-Minute .dat file
	echo -e "${DD}/${MN}/${YEAR}\t${HH}:${MM}\t${AvgPerMin}" >> ${file}_AvgPerMin.dat
	### IF statements to control incrementing of MM,HH,DD,MN,and YEAR
	if [ ${MM} -lt 59 ]
	then 
		let MM=${MM}+1 
	else 
		let MM=0
		if [ ${HH} -lt 23 ]
	       	then 
			let HH=${HH}+1
		else 
			let HH=0
			if [ ${MN} -eq 1 ]
		       	then 
				let days=31
			elif [ ${MN} -eq 2 ]
		       	then 
				let days=28
			elif [ ${MN} -eq 3 ]
		       	then 
				let days=31
			elif [ ${MN} -eq 4 ] 
			then 
				let days=30
			elif [ ${MN} -eq 5 ]
			then 
				let days=31
			elif [ ${MN} -eq 6 ]
			then 
				let days=30
			elif [ ${MN} -eq 7 ] 
			then 
				let days=31
			elif [ ${MN} -eq 8 ] 
			then 
				let days=31
			elif [ ${MN} -eq 9 ] 
			then 
				let days=30
			elif [ ${MN} -eq 10 ] 
			then 
				let days=31
			elif [ ${MN} -eq 11 ] 
			then 
				let days=30
			else 
				let days=31
			fi
			if [[ ${DD} -lt ${days} ]] 
			then 
				let DD=${DD}+1
			else 
				let DD=1
				if [ ${MN} -lt 12 ]
				then 
					let MN=${MN}+1
				else 
					let MN=1
					let YEAR=${YEAR}+1
				fi
			fi
		fi
	fi
done
### Convert the Average-Per-Minute .dat file to .xlsx file using ssconvert 
ssconvert ${file}_AvgPerMin.dat ${file}_AvgPerMin.xlsx
### Delete {file}.dat file
rm -f ${file}.dat 
### START COMPUTING AVERAGE PER HOUR
### first convert ${file}_AvgPerMin.xlsx back to .csv file
### then copy to a .dat file noting the file separation indicator
ssconvert ${file}_AvgPerMin.xlsx ${file}.csv
awk 'BEGIN{FS=","; OFS="\t";}{print $1,$2,$3}' ${file}.csv > ${file}.dat
### delete the first row having the column titles
sed -i 1d ${file}.dat
### Scan the values of DD,MN,YEAR,HH, and MN from the first row
let YEAR=$(sed -n 1p ${file}.dat | awk '{print substr($1,1,4)}' | bc)
let DD=$(sed -n 1p ${file}.dat | awk '{print substr($1,6,2)}' | bc)
let MN=$(sed -n 1p ${file}.dat | awk '{print substr($1,9,2)}' | bc)
let HH=$(sed -n 1p ${file}.dat | awk '{print substr($2,1,2)}' | bc)
let MM=$(sed -n 1p ${file}.dat | awk '{print substr($2,4,2)}' | bc)
### Compute average of the first hour
### but first determine the number of minutes in the first hour
let lines=60-${MM}
AvgPerHr=$(sed -n 1,${lines}p ${file}.dat | awk '{ total += $3; n++ } END { printf "%.2f",total/n}')
### Enter the column headings to the Average-Per-Hour .dat file
### Then append Date, HH, and AvgPerHr to the Average-Per-Hour .dat file
echo -e "Date\tHour\tAvg PT per hour (W)" > ${file}_AvgPerHr.dat
echo -e "${DD}/${MN}/${YEAR}\t${HH}:00\t${AvgPerHr}" >> ${file}_AvgPerHr.dat
### Having computed the average of the first hour, first update the next line to start computation
### Then Scan the values of DD,MN,YEAR, and HH from that line
let lines=${lines}+1
let YEAR=$(sed -n ${lines}p ${file}.dat | awk '{print substr($1,1,4)}' | bc)
let DD=$(sed -n ${lines}p ${file}.dat | awk '{print substr($1,6,2)}' | bc)
let MN=$(sed -n ${lines}p ${file}.dat | awk '{print substr($1,9,2)}' | bc)
let HH=$(sed -n ${lines}p ${file}.dat | awk '{print substr($2,1,2)}' | bc)
### Count the total_lines of ${file}.dat
wc -l ${file}.dat > temporaryfile.out
read total_lines filename < temporaryfile.out
rm -f temporaryfile.out
### FOR loop with the same concept as FOR loop above.
### only that we are considering MM not SS like in FOR loop above
for ((i=${lines}; i<=${total_lines}; i+=60));
do
	let lastminute=${i}+59
	AvgPerHr=$(sed -n ${i},${lastminute}p ${file}.dat | awk '{ total += $3; n++} END { printf "%.2f",total/n}')
	echo -e "${DD}/${MN}/${YEAR}\t${HH}:00\t${AvgPerHr}" >> ${file}_AvgPerHr.dat
	if [ ${HH} -lt 23 ] 
	then 
		let HH=${HH}+1
	else
		let HH=0
		if [ ${MN} -eq 1 ]
		then
			let days=31
		elif [ ${MN} -eq 2 ]
		then
			let days=28
		elif [ ${MN} -eq 3 ]
		then
			let days=31
		elif [ ${MN} -eq 4 ]
		then
			let days=30
		elif [ ${MN} -eq 5 ]
		then
			let days=31
		elif [ ${MN} -eq 6 ]
		then
			let days=30
		elif [ ${MN} -eq 7 ]
		then
			let days=31
		elif [ ${MN} -eq 8 ]
		then
			let days=31
		elif [ ${MN} -eq 9 ]
		then
			let days=30
		elif [ ${MN} -eq 10 ]
		then
			let days=31
		elif [ ${MN} -eq 11 ]
		then
			let days=30
		else
			let days=31
		fi
		if [[ ${DD} -lt ${days} ]]
		then
			let DD=${DD}+1
		else
			let DD=1
			if [ ${MN} -lt 12 ]
			then
				let MN=${MN}+1
			else
				let MN=1
				let YEAR=${YEAR}+1
			fi
		fi
	fi
done
### Convert Average-Per-Hour .dat file to .xlsx 
ssconvert ${file}_AvgPerHr.dat ${file}_AvgPerHr.xlsx
rm -f ${file}.dat ${file}.csv
