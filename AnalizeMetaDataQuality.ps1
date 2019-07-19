#===================================================================================================================#
#  Built and tested under Powershell version 5.1                                                                    #
#                                                                                                                   #
# This script imports a CSV UTF8 file ($CSV_inpfile), performing some Quality analysis upon its Contents and then   #
# saving results into an XLSX output file ($XLS_outfile).                                                           #
#                                                                                                                   #
# Main variables:                                                                                                   #
#                                                                                                                   #
#   Object[] $rows_from_csv:  contains all contents from the CSV input file                                         #
#   [hashtable] $data:        collects and arranges data from the Array $rows_from_csv to be analyzed               #
#   [hashtable] $stats:       This hashtable stores the results from the Analysis                                   #
#                                                                                                                   #
# file $CSV_inpfile description:  First row contains the headers for the 13 columns:                                #
#                                                                                                                   #
#  host_name,   host_short_desc,     host_long_desc,                                                                #
#  database,    database_short_desc, database_long_desc,                                                            #
#  schema_name, schema_short_desc,   schema_long_desc,                                                              #
#  table_name,  table_short_desc,    table_long_desc,                                                               #
#  col_name,    col_short_desc,      col_long_desc                                                                  #
#                                                                                                                   #
#================================================================================================================== #
param (
       [Parameter(mandatory=$false, Position=0, HelpMessage="Please select your input CSV file.")] [string] $inp,
       [Parameter(mandatory=$false, Position=0)] [string] $out,
       [Parameter(mandatory=$false, Position=0)] [string] $fullReport='N'
)

Set-StrictMode -Version 1.0


# Get this Script Name 
$ThisScript = (Get-Item $PSCommandPath ).Name 

# Check input parameters
if ( ! $inp  ){
   echo "===================================================================================================================================================================="
   echo "Missing input parameter!"
   echo " "
   echo "   Usage: .\$ThisScript -inp <Input_CSV_File>  [-out <output_directory>]  [-fullReport <Y/N>]"
   echo " "
   echo "Examples: "
   echo "          .\$ThisScript -inp .\input_CSV_files\IDW.csv                           | In this case the output Excel file is created in current directory "
   echo "          .\$ThisScript -inp .\input_CSV_files\IDW.csv  -out output_XLSX_files   | In this case the output Excel file is created in directory 'output_XLSX_files'   "
   echo "          .\$ThisScript -inp .\input_CSV_files\IDW.csv  -fullReport Y            | In this case the output Excel file is created with all Sheets (empty or not)     "
   echo "===================================================================================================================================================================="
   echo " "
   exit
}

# set date and time to use for output filename
$datetime=Get-Date
$datetime_str = "" + $datetime.Year + 
                "-" + $datetime.Month.ToString().PadLeft(2,'0')  + 
                "-" + $datetime.Day.ToString().PadLeft(2,'0')    + 
                "_" + $datetime.Hour.ToString().PadLeft(2,'0')   + 
                "h" + $datetime.Minute.ToString().PadLeft(2,'0')


# ============================================ #
#                                              #
# Setting some analysis criteria parameters    #
#                                              #
# ============================================ #

$MIN_HOST_DESC_NCHARS_PER_WORD = 3
$MIN_DB_DESC_NCHARS_PER_WORD   = 3
$MIN_SCHE_DESC_NCHARS_PER_WORD = 3
$MIN_TAB_DESC_NCHARS_PER_WORD  = 3
$MIN_COL_DESC_NCHARS_PER_WORD  = 3

$MIN_HOST_DESC_NWORDS = 2
$MIN_DB_DESC_NWORDS   = 2
$MIN_SCHE_DESC_NWORDS = 2
$MIN_TAB_DESC_NWORDS  = 2
$MIN_COL_DESC_NWORDS  = 2

$host_short_desc_allowed = @();
$db_short_desc_allowed = @();
$schema_short_desc_allowed = @();
$tab_short_desc_allowed=@();
$col_short_desc_allowed=@();

$host_short_desc_forbidden = @();
$db_short_desc_forbidden = @();
$schema_short_desc_forbidden = @();
$tab_short_desc_forbidden=@("description");
$col_short_desc_forbidden=@("description");

<#  A regexp $split_regexp é usada para fazer o split das descrições em palavras. O split é feito por sequências dos tipos seguintes:

    - 0 ou mais espacos, seguido de um dos caracteres seguintes - [&,:()/.]  -  seguido novamente de zero ou mais espaços
    - 1 ou mais espacos     
#>
$split_regexp=[regex]"\s*[;,&\:\(\)/\.=\*'?]+\s*|\s+"

# this regexp detects characters outside bellow range
$chars_notallowed_inside_a_word=[regex]"[^a-zA-Z0-9\-ãÃáÁàÀâÂéÉêÊíÍóÓõÕúÚçÇ´‘’_]"


# =============================================================================== #
#                                                                                 #
# processing parameters to set output filename (based on name from input file)    #
#                                                                                 #
# =============================================================================== #

$baseName=(Get-Item $inp).Name -replace ".csv", ""
$baseInputDir=(Get-Item $inp).DirectoryName

if ( $out -and (Test-Path $out) -and (  (Get-Item $out).Attributes -eq "Directory" )){
    $baseOutputDir=(Get-Item $out).FullName
}else{
    $baseOutputDir=(Get-Item $PSCommandPath ).DirectoryName
}

$CSV_inpfile = $baseInputDir+"\"+$baseName+".csv"
if ( $fullReport -eq 'Y'){
    $XLS_outfile = $baseOutputDir+"\"+$baseName+"_analyzed_"+$datetime_str+"-allSheets.xlsx"
}else{
    $XLS_outfile = $baseOutputDir+"\"+$baseName+"_analyzed_"+$datetime_str+".xlsx"
}

# =============================================================================== #
#                                                                                 #
# ======                FUNCTIONS DECLARATION SECTION                      ====== #
#                                                                                 #
# =============================================================================== #


function StartStop-Print ( $local_step, $global_step, $text_input, $time0) {

  $text = $global_step.ToString();
  $text = $text.PadLeft(2,'0');

  $time=Get-Date; 


  If ( $local_step -eq "1" ) {

    $text += ") Start - " + $text_input

    $text += " at "+ $time

  }Else{

    $text += ")   End - " + $text_input

    $text += " at "+ $time

    If ( $time0 ){
        $time0 = [DateTime] $time0;
        $dt = $time - $time0;
        $dt_min = $dt.Minutes;
        $dt_sec = $dt.Seconds;
        $dt_msec = $dt.Milliseconds;

        If ( $dt_min -gt 1){
            $text += " (exec. time: $dt_min min $dt_sec sec.) "
        }Elseif($dt_sec -gt 1){
            $text += " (exec. time: $dt_sec sec.) "
        }Else{
            $text += " (exec. time: $dt_msec msec.) "
        }
    }
    $text += "`r`n"
  }
  
  
  If ( $local_step -eq "1" ) {
        Write-Host ""
  }
  Write-Host $text -ForegroundColor Yellow # -BackgroundColor DarkYellow
  
  return $time
}

<#

Function to add worksheet and fill the list of values from input data.
---------------------------------------------------------------------

It receives folowing inputs:

 - Object: reference to the active workbook
 - String: worksheet Name (use number to have a uniform nomenclature)
 - String: Title to the worksheet
 - Object: input data

#>

Function addWorkSheet($workbook, $sheetName, $title, $headers, $inputData , $freq){

    # >>> if there is no data, simply leave and do not add any sheet   
    if ($inputData.Count -eq 0 -and $fullReport -ne 'Y'){

        return

    }

    # Add a workSheet after the last one
    $workSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $workbook.Worksheets.Item($Workbook.Worksheets.Count)) 

    # add Name
    $workSheet.Name = "Sheet "+ $sheetName
    
    # add link in first cell 
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item(1,1), "", "'Stats'!A1", "Ir para a sheet Stats", "Stats") | Out-Null

    # add title in second cell
    $workSheet.Cells.Item(2,2) = $title

    # merge cells from title (by merging all second row)
    #$workSheet.Rows("2").MergeCells = $true 
    $workSheet.Cells.Range("B2:L2").MergeCells = $true
 
    # >>> if there is no data, write it down and leave   
    if ($inputData.Count -eq 0){
        
        $workSheet.Cells.Item(3,2) = "Não existem registos para este critério"
        $workSheet.Rows("3").MergeCells = $true 
        return

    }
 
    # >>> In case there is data...
        
    $ncols = $headers.Count 
    $nrows = $inputData.Count
    $rows =  new-object 'Object[,]' $nrows, $ncols

    if ( $freq ){
        $last_col_index = $ncols-1
        
        $words = $title -split [regex]'\s+' -match '\S+'

        if ( $words[0] -eq 'Colunas'){
            $ref_col_index = 5
        }elseif( $words[0] -eq 'Tabelas'){
            $ref_col_index = 3
        }elseif( $words[0] -eq 'Schemas'){
            $ref_col_index = 2
        }elseif( $words[0] -eq 'BD'){
            $ref_col_index = 1
        }
    }

    # set start of the range to use
    $row_headers = 3
    $first_col   = 2
    $last_col    = $first_col   + $ncols  - 1
    #
    $first_row = $row_headers + 1
    $last_row  = $row_headers + $nrows
    
    $dataRange = $workSheet.Range($workSheet.Cells($first_row, $first_col), $workSheet.Cells($last_row, $last_col) )

    # >>> set headers
    for ( $i = 0; $i -lt $ncols; $i++){        
        $workSheet.Cells.Item($row_headers,$first_col+$i) = $headers[$i]        
    }
    # >>> copy data from $inputData list into array $rows     
    for ($i = 0; $i -lt $nrows; $i++){
        for ($j = 0; $j -lt $ncols; $j++){
            $rows[$i,$j] = $inputData[$i][$j]
        }
        # if there is frequency data, rewrite $last element
        if ( $freq ){
            $key = $inputData[$i][$ref_col_index]
            $rows[$i,$last_col_index] = $freq[$key]
        }
    }
    # copy data into the Excel object (in a very fastet way than writing to $workSheet.Cells like in the case of the headers
    $dataRange.Value2 = $rows


    # >>> Set some presentation settings

    # set Bold font and autofilters on the headers range

    #$headersRange = "B3:G3"
    #$workSheet.Cells.Range($headersRange).Font.Bold=$True
    #$workSheet.Cells.Range($headersRange).Columns.AutoFilter() | Out-Null
    
    $headersRange = $workSheet.Range($workSheet.Cells($row_headers, $first_col), $workSheet.Cells($row_headers, $last_col) )
    $headersRange.Font.Bold=$True
    $headersRange.Columns.AutoFilter() | Out-Null

    # Set automatic width for all columns, adjust for vertical centered alignement, and other stuff...
    $usedRange = $workSheet.UsedRange()
    $usedRange.EntireColumn.AutoFit() | Out-Null
    $usedRange.Rows.VerticalAlignment = -4108
    $usedRange.Rows.rowHeight = 18

    $WorkSheet.Application.ActiveWindow.SplitColumn = 1
    $WorkSheet.Application.ActiveWindow.SplitRow = 3
    $WorkSheet.Application.ActiveWindow.FreezePanes = $true


}

Function analyzeMetadata($in, $out, $db_host, $db, $schema, $table, $col ){

	# detect level of analysis
	if ( $col ){
		$level = 5
        $field_name = $col
        $tab_short_desc = $in[$db_host][$db][$schema][$table]["metadata"]["short_desc"].trim()		
        #
        $short_desc_nwords_min = $MIN_COL_DESC_NWORDS
		$short_desc_nchars_per_word_min = $MIN_COL_DESC_NCHARS_PER_WORD
        $short_desc_forbidden = $col_short_desc_forbidden
        $short_desc_allowed   = $col_short_desc_allowed

	}elseif ( $table){
		$level = 4
        $field_name = $table
		$short_desc_nwords_min = $MIN_TAB_DESC_NWORDS
		$short_desc_nchars_per_word_min = $MIN_TAB_DESC_NCHARS_PER_WORD
        $short_desc_forbidden = $tab_short_desc_forbidden
        $short_desc_allowed   = $tab_short_desc_allowed
	}elseif ( $schema){
		$level = 3
        $field_name = $schema
		$short_desc_nwords_min = $MIN_SCHE_DESC_NWORDS
		$short_desc_nchars_per_word_min = $MIN_SCHE_DESC_NCHARS_PER_WORD
        $short_desc_forbidden = $schema_short_desc_forbidden
        $short_desc_allowed   = $schema_short_desc_allowed
	}elseif ( $db){
		$level = 2
        $field_name = $db
		$short_desc_nwords_min = $MIN_DB_DESC_NWORDS
		$short_desc_nchars_per_word_min = $MIN_DB_DESC_NCHARS_PER_WORD
        $short_desc_forbidden = $db_short_desc_forbidden
        $short_desc_allowed   = $db_short_desc_allowed
	}elseif ( $db_host){
		$level = 1
        $field_name = $db_host
		$short_desc_nwords_min = $MIN_HOST_DESC_NWORDS
		$short_desc_nchars_per_word_min = $MIN_HOST_DESC_NCHARS_PER_WORD
        $short_desc_forbidden = $host_short_desc_forbidden
        $short_desc_allowed   = $host_short_desc_allowed
	}

	# Seek for the metadata at the right level
	if ( $level -eq 5 ){
		$short_desc = $in[$db_host][$db][$schema][$table][$col]["metadata"]["short_desc"].trim()

        # don't understand this exception but it's needed
        if ($data[$db_host][$db][$schema][$table][$col]["metadata"]["long_desc"]){
            $long_desc = $data[$db_host][$db][$schema][$table][$col]["metadata"]["long_desc"].trim()
        }else{
            $long_desc = ""  
        }

	}elseif ( $level -eq 4 ){
		$short_desc = $in[$db_host][$db][$schema][$table]["metadata"]["short_desc"].trim()
		$long_desc = $in[$db_host][$db][$schema][$table]["metadata"]["long_desc"].trim()
	}elseif ( $level -eq 3 ){
		$short_desc = $in[$db_host][$db][$schema]["metadata"]["short_desc"].trim()
		$long_desc = $in[$db_host][$db][$schema]["metadata"]["long_desc"].trim()
	}elseif ( $level -eq 2 ){
		$short_desc = $in[$db_host][$db]["metadata"]["short_desc"].trim()
		$long_desc = $in[$db_host][$db]["metadata"]["long_desc"].trim()
	}elseif ( $level -eq 1 ){
		$short_desc = $in[$db_host]["metadata"]["short_desc"].trim()
		$long_desc = $in[$db_host]["metadata"]["long_desc"].trim()
	}

	# Metric: Count all Objects analyzed
	if ( $level -eq 5 ){
		$out["Colunas"].Add(@($db, $schema, $table, $col))  
	}elseif ( $level -eq 4 ){
		$out["Tabelas"].Add(@($db, $schema, $table))  
	}elseif ( $level -eq 3 ){
		$out["Schemas"].Add(@($db, $schema))  
	}elseif ( $level -eq 2 ){
		$out["BD"].Add(@($db))  
	}elseif ( $level -eq 1 ){
		$out["Host"].Add(@($db_host))  	
	}
		


	# Objects without short description
	If ($short_desc.equals("") -or $short_desc.ToUpper().equals("NULL") ){

		# Metric: No short Neither Long Description 
		If ( $long_desc -and (! $long_desc.equals("") -and ! $long_desc.ToUpper().equals("NULL") ) ){

			if ( $level -eq 5 ){
				$out["Colunas com Long Desc mas sem short Desc"].Add(@($db, $schema, $table, $tab_short_desc, $col))
			}elseif ( $level -eq 4 ){
				$out["Tabelas com Long Desc mas sem short Desc"].Add(@($db, $schema, $table))
			}elseif ( $level -eq 3 ){
				$out["Schemas com Long Desc mas sem short Desc"].Add(@($db, $schema))
			}elseif ( $level -eq 2 ){
				$out["BD com Long Desc mas sem short Desc"].Add(@($db))
			}elseif ( $level -eq 1 ){
				$out["Host com Long Desc mas sem short Desc"].Add(@($db_host))
			}
		# Metric: There is no short Desc but there is a Long Description!
		}Else{
			if ( $level -eq 5 ){
				$out["Colunas sem short Desc nem Long Desc"].Add(@($db, $schema, $table, $tab_short_desc, $col))
			}elseif ( $level -eq 4 ){
				$out["Tabelas sem short Desc nem Long Desc"].Add(@($db, $schema, $table))
			}elseif ( $level -eq 3 ){
				$out["Schemas sem short Desc nem Long Desc"].Add(@($db, $schema))
			}elseif ( $level -eq 2 ){
				$out["BD sem short Desc nem Long Desc"].Add(@($db))
			}elseif ( $level -eq 1 ){
				$out["Host sem short Desc nem Long Desc"].Add(@($db_host))
			}						
		}

	}


	# ==> Analyze Eligible short descriptions (meaning not NULL valued)
	If (  $short_desc -and $short_desc.Length -gt 0 -and ! $short_desc.ToUpper().equals("NULL") ){                                 
		
		# init counter vars
		$single_word_desc = $null; $nw=0; $nchars=0 ; $nchars_max = 0; $numb_special_chars = 0                   

		# split the description into words and apply some analysis criteria
		$words = $short_desc -Split $split_regexp -match '\S+'
		foreach ($word in $words){

			# init boolean used to classify the candidate word
			$its_a_valid_word = $true

			# Apply 1st rule: Valid words do not start by decimal numbers characters
			if ( $word -match '^\d+' ){
				
				$its_a_valid_word = $false
			} 

			# Apply 2nd rule:  Valid Words must have at least $short_desc_nchars_per_word_min characters
			if ( $word.Length -lt $short_desc_nchars_per_word_min ){
				
				$its_a_valid_word = $false
			}

			# Apply 3rd rule: Test if the candidate word belongs to a list of acceptable words (even if not complying with previous rules)
			foreach ($pattern in $short_desc_allowed){                                
					if ($word -imatch $pattern){

						$its_a_valid_word = $true
					}                                
			}

			# Candidate has passed all tests, so process it
			if ( $its_a_valid_word ) {
				$nw++
				$nchars += $word.Length

				$single_word_desc = $word

				if ( $word.Length -gt $nchars_max ){                    
					$nchars_max = $word.Length
				}
				# detect non-ASCII characters
				if( $word -match $chars_notallowed_inside_a_word){
					$numb_special_chars++
				}                            
			}

		}
		# compute average number of characters per word
		$nchar_avg = 0
		if ( $nw -gt 0){
			$nchar_avg = [math]::Round($nchars/$nw,1)
		}    		
		if ( $level -eq 5 ){
			$out["Colunas com short Desc"].Add( @($db, $schema, $table, $tab_short_desc, $col, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max  ) )
		}elseif ( $level -eq 4 ){
			$out["Tabelas com short Desc"].Add( @($db, $schema, $col, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max  ) )
		}elseif ( $level -eq 3 ){
			$out["Schemas com short Desc"].Add( @($db, $schema, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max  ) )
		}elseif ( $level -eq 2 ){
			$out["BD com short Desc"].Add( @($db, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max  ) )
		}elseif ( $level -eq 1 ){
			$out["Host com short Desc"].Add( @($db_host, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max  ) )
		}	
		
		
		# Description has $nw words with a total of $nchars characters (without account of the single characters dropped on the split regexp - blanks and others...)


		# if the description is single-worded, make aditional tests 
		# ==========================================================

		if ( $nw -eq 1){

            # Test if the candidate word belongs to a list of forbidden words
			foreach ($pattern in $short_desc_forbidden){
				if ($single_word_desc -imatch $pattern){
					$nw = 0
				}   
			}
            # reject a description which contains the name of the field being analyzed
            if ( $single_word_desc -imatch $field_name){
                $nw = 0
            }

		}

		# Now, we apply the cut-off criteria to the surviving descriptions and populate the corresponding warning lists
		# =============================================================================================================

		# we have at least one Word!
		IF ( $nw -ne 0 ){

                                                     
			# Column Metric: Too short Description (in numb words)
            If ( $nw -eq 1 ){

                $valid_desc = $false
			    # test if the single-worded description belongs to a list of acceptable words
			    foreach ($pattern in $short_desc_allowed){                                
					    if ($single_word_desc -imatch $pattern){
                            $valid_desc = $true
					    }                                
			    }
                
                if ( $valid_desc -eq $false ){
				    if ( $level -eq 5 ){
					    $out["Desc. Coluna inválida (em nº palavras)"].Add(@($db, $schema, $table, $tab_short_desc, $col, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
				    }elseif ( $level -eq 4 ){
					    $out["Desc. Tabela inválida (em nº palavras)"].Add(@($db, $schema, $table, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
				    }elseif ( $level -eq 3 ){
					    $out["Desc. Schema inválida (em nº palavras)"].Add(@($db, $schema, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
				    }elseif ( $level -eq 2 ){
					    $out["Desc. BD inválida (em nº palavras)"].Add(@($db, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
				    }elseif ( $level -eq 1 ){
					    $out["Desc. Host inválida (em nº palavras)"].Add(@($db_host, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
				    }                                
                }


			}elseif ( $nw -lt $short_desc_nwords_min ){


				if ( $level -eq 5 ){
					$out["Desc. Coluna inválida (em nº palavras)"].Add(@($db, $schema, $table, $tab_short_desc, $col, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
				}elseif ( $level -eq 4 ){
					$out["Desc. Tabela inválida (em nº palavras)"].Add(@($db, $schema, $table, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
				}elseif ( $level -eq 3 ){
					$out["Desc. Schema inválida (em nº palavras)"].Add(@($db, $schema, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
				}elseif ( $level -eq 2 ){
					$out["Desc. BD inválida (em nº palavras)"].Add(@($db, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
				}elseif ( $level -eq 1 ){
					$out["Desc. Host inválida (em nº palavras)"].Add(@($db_host, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
				}
			}

		
		# we don't have valid words but we have something that has to be traced
		}ELSE{

			# Metric: short desc contains invalid description (the candidate words are not real words...)
			if ( $level -eq 5 ){
				$out["Desc. Coluna inválida (em nº caract./palavra)"].Add(@($db, $schema, $table, $tab_short_desc, $col, $short_desc, $long_desc, $nw, $nchars, $nchars_max))
			}elseif ( $level -eq 4 ){
				$out["Desc. Tabela inválida (em nº caract./palavra)"].Add(@($db, $schema, $table, $short_desc, $long_desc, $nw, $nchars, $nchars_max))
			}elseif ( $level -eq 3 ){
				$out["Desc. Schema inválida (em nº caract./palavra)"].Add(@($db, $schema, $short_desc, $long_desc, $nw, $nchars, $nchars_max))
			}elseif ( $level -eq 2 ){
				$out["Desc. BD inválida (em nº caract./palavra)"].Add(@($db, $short_desc, $long_desc, $nw, $nchars, $nchars_max))
			}elseif ( $level -eq 1 ){
				$out["Desc. Host inválida (em nº caract./palavra)"].Add(@($db_host, $short_desc, $long_desc, $nw, $nchars, $nchars_max))
			}


		}

		# Metric: short desc contains non-ASCII characters
		if ( $numb_special_chars -gt 0){

			if ( $level -eq 5 ){
				$out["Desc. Coluna contém caracteres 'especiais'"].Add(@($db, $schema, $table, $col, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
			}elseif ( $level -eq 4 ){
				$out["Desc. Tabela contém caracteres 'especiais'"].Add(@($db, $schema, $table, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
			}elseif ( $level -eq 3 ){
				$out["Desc. Schema contém caracteres 'especiais'"].Add(@($db, $schema, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
			}elseif ( $level -eq 2 ){
				$out["Desc. BD contém caracteres 'especiais'"].Add(@($db, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
			}elseif ( $level -eq 1 ){
				$out["Desc. Host contém caracteres 'especiais'"].Add(@($db_host, $short_desc, $long_desc, $nw, $nchars, $nchar_avg, $nchars_max))
			}
		}

		#  Column Short desc. freq.
		if ( $level -eq 5 ){
			if ( ! $out["Frequencias"]["Col. short desc."].ContainsKey($short_desc)){
				$out["Frequencias"]["Col. short desc."].Add($short_desc, 1)
			}else{
				$out["Frequencias"]["Col. short desc."][$short_desc] += 1
			}
		}elseif ( $level -eq 4 ){
			if ( ! $out["Frequencias"]["Tab. short desc."].ContainsKey($short_desc)){
				$out["Frequencias"]["Tab. short desc."].Add($short_desc, 1)
			}else{
				$out["Frequencias"]["Tab. short desc."][$short_desc] += 1
			}
		}elseif ( $level -eq 3 ){
			if ( ! $out["Frequencias"]["Schema short desc."].ContainsKey($short_desc)){
				$out["Frequencias"]["Schema short desc."].Add($short_desc, 1)
			}else{
				$out["Frequencias"]["Schema short desc."][$short_desc] += 1
			}
		}elseif ( $level -eq 2 ){
			if ( ! $out["Frequencias"]["BD short desc."].ContainsKey($short_desc)){
				$out["Frequencias"]["BD short desc."].Add($short_desc, 1)
			}else{
				$out["Frequencias"]["BD short desc."][$short_desc] += 1
			}
#		}elseif ( $level -eq 1 ){

		}




	} # End if ==> Analyze "Eligible short descriptions"

}


# =============================================================================== #
#                                                                                 #
# ======                SCRIPT BEGINS HERE                                 ====== #
#                                                                                 #
# =============================================================================== #


# init Excel output object
$excel = New-Object -ComObject excel.application
$excel.visible = $False
$excel.DisplayAlerts = $false #Supress alert messages.


# First Step
$step=1;
$input="import of CSV file $baseName"+".csv"
$time = StartStop-Print "1" $step $input;

$time0 = $time

$rows_from_csv = Import-Csv $CSV_inpfile -Encoding UTF8 -Delimiter "," 
$nb_rows_from_csv = $rows_from_csv.Count


$time = StartStop-Print "2" $step $input $time;


# init the $data hashtable - structure will be hastable of hashtable of list:  $schemas => $tables => $columns
#$global:data=@{};  
$data=@{};  

# init the $stats hastable
# $global:stats = @{
$stats = @{

    "Host" = New-Object System.Collections.Generic.List[System.Object];
    "Host sem short Desc nem Long Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Host com Long Desc mas sem short Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Host com short Desc igual ao Nome Coluna" = New-Object System.Collections.Generic.List[System.Object];
    "Host com short Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. Host contém caracteres 'especiais'" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. Host inválida (em nº palavras)" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. Host inválida (em nº caract./palavra)" = New-Object System.Collections.Generic.List[System.Object];

    "BD" = New-Object System.Collections.Generic.List[System.Object];
    "BD sem short Desc nem Long Desc" = New-Object System.Collections.Generic.List[System.Object];
    "BD com Long Desc mas sem short Desc" = New-Object System.Collections.Generic.List[System.Object];
    "BD com short Desc igual ao Nome Coluna" = New-Object System.Collections.Generic.List[System.Object];
    "BD com short Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. BD contém caracteres 'especiais'" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. BD inválida (em nº palavras)" = New-Object System.Collections.Generic.List[System.Object];    
    "Desc. BD inválida (em nº caract./palavra)" = New-Object System.Collections.Generic.List[System.Object];

    "Schemas" = New-Object System.Collections.Generic.List[System.Object];
    "Schemas sem short Desc nem Long Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Schemas com Long Desc mas sem short Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Schemas com short Desc igual ao Nome Coluna" = New-Object System.Collections.Generic.List[System.Object];
    "Schemas com short Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. Schema contém caracteres 'especiais'" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. Schema inválida (em nº palavras)" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. Schema inválida (em nº caract./palavra)" = New-Object System.Collections.Generic.List[System.Object];

    
    "Tabelas" = New-Object System.Collections.Generic.List[System.Object];
    "Tabelas sem short Desc nem Long Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Tabelas com Long Desc mas sem short Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Tabelas com short Desc igual ao Nome Coluna" = New-Object System.Collections.Generic.List[System.Object];
    "Tabelas com short Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. Tabela contém caracteres 'especiais'" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. Tabela inválida (em nº palavras)" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. Tabela inválida (em nº caract./palavra)" = New-Object System.Collections.Generic.List[System.Object];

    
    "Colunas" = New-Object System.Collections.Generic.List[System.Object];
    "Colunas sem short Desc nem Long Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Colunas com Long Desc mas sem short Desc" = New-Object System.Collections.Generic.List[System.Object];
    "Colunas com short Desc igual ao Nome Coluna" = New-Object System.Collections.Generic.List[System.Object];
    "Colunas com short Desc" = New-Object System.Collections.Generic.List[System.Object];    
    "Desc. Coluna contém caracteres 'especiais'" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. Coluna inválida (em nº palavras)" = New-Object System.Collections.Generic.List[System.Object];
    "Desc. Coluna inválida (em nº caract./palavra)" = New-Object System.Collections.Generic.List[System.Object];

    "Frequencias" = @{ 
        "BD short desc." =  @{}; 
        "Schema short desc." =  @{}; 
        "Tab. short desc." =  @{}; 
        "Col. short desc." =  @{}; 
    }
};


$step++; 
$input = "loading $nb_rows_from_csv rows from file $baseName"+".csv into data structure"+ ' $data';
$time = StartStop-Print "1" $step $input $time;

$nrows=0;
ForEach ($row in $rows_from_csv){
    $nrows++
	
    ### Read contents from row into local variables

    $db_host = $row.('host_name');
    $db_host_short_desc = $row.('host_short_desc');
    $db_host_long_desc = $row.('host_long_desc');

    $db = $row.('database');
    $db_short_desc = $row.('database_short_desc');
    $db_long_desc = $row.('database_long_desc');

    $schema = $row.('schema_name');
    $schema_short_desc = $row.('schema_short_desc');
    $schema_long_desc = $row.('schema_long_desc');


    $table = $row.('table_name');
    $table_short_desc = $row.('table_short_desc');
    $table_long_desc = $row.('table_long_desc');
    

    $col = $row.('col_name');
    $col_short_desc = $row.('col_short_desc');
    $col_long_desc = $row.('col_long_des');


	####  Fill entries from super hastable $data  
    
    If ( ! $data.ContainsKey($db_host) ) {
        
        $data.Add($db_host, @{ 'metadata'= @{ 'short_desc'= $db_host_short_desc;  'long_desc' = $db_host_long_desc} })
        $data[$db_host].Add($db, @{ 'metadata'= @{ 'short_desc'= $db_short_desc;  'long_desc' = $db_long_desc} })
        $data[$db_host][$db].Add($schema, @{ 'metadata'= @{ 'short_desc'= $schema_short_desc;  'long_desc' = $schema_long_desc} })
        $data[$db_host][$db][$schema].Add($table , @{ 'metadata' = @{'short_desc'= $table_short_desc;  'long_desc' = $table_long_desc} } )
        $data[$db_host][$db][$schema][$table].Add($col , @{ 'metadata'= @{'short_desc'= $col_short_desc;  'long_desc' = $col_long_desc} } )

    }elseif ( ! $data[$db_host].ContainsKey($db) ) {

        $data[$db_host].Add($db, @{ 'metadata'= @{ 'short_desc'= $db_short_desc;  'long_desc' = $db_long_desc} })
        $data[$db_host][$db].Add($schema, @{ 'metadata'= @{ 'short_desc'= $schema_short_desc;  'long_desc' = $schema_long_desc} })
        $data[$db_host][$db][$schema].Add($table , @{ 'metadata' = @{'short_desc'= $table_short_desc;  'long_desc' = $table_long_desc} } )
        $data[$db_host][$db][$schema][$table].Add($col , @{ 'metadata'= @{'short_desc'= $col_short_desc;  'long_desc' = $col_long_desc} } )

    }elseif ( ! $data[$db_host][$db].ContainsKey($schema) ) {
        
        $data[$db_host][$db].Add($schema, @{ 'metadata'= @{ 'short_desc'= $schema_short_desc;  'long_desc' = $schema_long_desc} })
        $data[$db_host][$db][$schema].Add($table , @{ 'metadata' = @{'short_desc'= $table_short_desc;  'long_desc' = $table_long_desc} } )
        $data[$db_host][$db][$schema][$table].Add($col , @{ 'metadata'= @{'short_desc'= $col_short_desc;  'long_desc' = $col_long_desc} } )

    }elseif ( ! $data[$db_host][$db][$schema].ContainsKey($table) ) {

        $data[$db_host][$db][$schema].Add($table , @{ 'metadata' = @{'short_desc'= $table_short_desc;  'long_desc' = $table_long_desc} } )
        $data[$db_host][$db][$schema][$table].Add($col , @{ 'metadata'= @{'short_desc'= $col_short_desc;  'long_desc' = $col_long_desc} } )

    }elseif ( ! $data[$db_host][$db][$schema][$table].ContainsKey($col) ) {

        $data[$db_host][$db][$schema][$table].Add($col , @{ 'metadata'= @{'short_desc'= $col_short_desc;  'long_desc' = $col_long_desc} } )

    }
    

}
$time = StartStop-Print "2" $step $input $time;

#Remove-Variable -Name $rows_from_csv
$rows_from_csv = $null

#---------------------------------------------------------------------------------------------------------------------#
#                                                                                                                     #
# Filling stats information into Hastable $stats:                                                                     #
#                                                                                                                     #
# $stats[metric] = { @{ $db , $schema, $table, $col }}                                                                #
#                                                                                                                     #
# Metric = Métricas analisadas:  Ver a inicialização do hashtable $stats para conferir quais as métricas analisadas   #
#                                                                                                                     #
#---------------------------------------------------------------------------------------------------------------------#



#=================================================================================#
# Perform the Analysis upon hastable $data and store results into hastable $stats #
#=================================================================================#

$step++; 
$input = "Analyzing data from hastable "+'$data';
$time = StartStop-Print "1" $step $input $time;


foreach( $db_host in $data.keys){

    analyzeMetadata -in $data -out $stats -db_host $db_host


    foreach( $key in $data[$db_host].keys){
    #=======================================#

        # Skip key "metadata"
        If ( $key -eq "metadata") {            
            continue
        }
        $db = $key

        analyzeMetadata -in $data -out $stats -db_host $db_host -db $db
     

        foreach ( $key in $data[$db_host][$db].keys){
        #=============================================#

            # Skip key "metadata"
            If ( $key -eq "metadata") {            
                continue
            }
            $schema = $key

            analyzeMetadata -in $data -out $stats -db_host $db_host -db $db -schema $schema


            foreach ( $key in $data[$db_host][$db][$schema].keys){
            #======================================================#
        
                # Skip key "metadata"
                If ( $key -eq "metadata") {            
                    continue
                }
                $table = $key

                analyzeMetadata -in $data -out $stats -db_host $db_host -db $db -schema $schema -table $table



                foreach ( $key in $data[$db_host][$db][$schema][$table].keys){ 
                #===============================================================#
                    
                    # Skip key "metadata"
                    If ( $key -eq "metadata") {            
                        continue
                    }
                    $col = $key

                    analyzeMetadata -in $data -out $stats -db_host $db_host -db $db -schema $schema -table $table -col $col


                } # end of loop on columns                  --> loop on $data[$db_host][$db][$schema][$table]

            } # End of loop on tables                       --> loop on $data[$db_host][$db][$schema]
        
        } # End of loop on schemas                          --> loop on $data[$db_host][$db]

    } # End of the loop on databases                        --> loop on $data[$db_host] 
    
}# End of the loop on hosts                                 --> loop on $data (Outer Analysis loop)

$time = StartStop-Print "2" $step $input $time;

$data = $null
#Remove-Variable -Name $data


# ===================================================#
#                                                    #
# Writing results from $stats into Output Excel file #
#                                                    #
# ===================================================#

# Add a new Empty Workbook  
$global:workbook = $excel.Workbooks.Add()   # this gives us one worksheet (named by default "sheet 1" and used at the end of script to fill the Sheet "Regras")

$ws_nb = 0;

############## Sheet "Host sem short Desc nem Long Desc"

$step++ 
$ws_nb++
$title = "Host sem short Desc. nem Long Desc. preenchidas"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;


$headers = @("Host", "Short Desc.", "Long Desc.")

addWorkSheet -workbook $workbook -sheetName '01' -title $title -headers $headers -inputData $stats["Host sem short Desc nem Long Desc"]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "BD sem short Desc nem Long Desc"

$step++ 
$ws_nb++
$title = "BD sem short Desc. nem Long Desc. preenchidas"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;


$headers = @("BD", "Short Desc.", "Long Desc.")

addWorkSheet -workbook $workbook -sheetName '11' -title $title -headers $headers -inputData $stats["BD sem short Desc nem Long Desc"]

$time = StartStop-Print "2" $step $input $time;

############## Sheet "BD sem short Desc nem Long Desc"

$step++ 
$ws_nb++
$title = "BD com Long Desc. preenchida mas sem short Desc."
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;


$headers = @("BD", "Short Desc.", "Long Desc.")

addWorkSheet -workbook $workbook -sheetName '12' -title $title  -headers $headers -inputData $stats["BD com Long Desc mas sem short Desc"]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "BD com short Desc"

$step++ 
$ws_nb++
$title = "BD com short Desc. preenchida"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;


$headers = @("BD", "Short Desc.", "Long Desc.",
             "Nº palavras", "Nº caract. Short Desc.", 
             "Nº caract./palavra (média)", "Nº caract./palavra (máximo)", 
             "Frequência"
)

addWorkSheet -workbook $workbook -sheetName '13' -title $title  -headers $headers -inputData $stats["BD com short Desc"] # -freq $stats["Frequencias"]["BD short desc."]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "BD com Short Desc. inválida (em nº caract./palavra)"

$step++ 
$ws_nb++
$title = "BD com Short Desc. inválida (em nº caract./palavra)"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;


$headers = @("BD", "Short Desc.", "Long Desc.",
             "Frequência")

addWorkSheet -workbook $workbook -sheetName '14' -title $title  -headers $headers -inputData $stats["Desc. BD inválida (em nº caract./palavra)"] # -freq $stats["Frequencias"]["BD short desc."]

$time = StartStop-Print "2" $step $input $time;

############## Sheet "BD com Short Desc. inválida (em nº palavras)"

$step++ 
$ws_nb++
$title = "BD com Short Desc. inválida (em nº palavras)"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;


$headers = @("BD", "Short Desc.", "Long Desc.",
             "Nº palavras", "Nº caract. Short Desc.", 
             "Nº caract./palavra (média)", "Nº caract./palavra (máximo)", 
             "Frequência")

addWorkSheet -workbook $workbook -sheetName '15' -title $title  -headers $headers -inputData $stats["Desc. BD inválida (em nº palavras)"] # -freq $stats["Frequencias"]["BD short desc."]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "Desc. BD contém caracteres 'especiais'"

$step++ 
$ws_nb++
$title = "BD - Short Desc. contém caracteres 'especiais'"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;


$headers = @("BD", "Short Desc.", "Long Desc.",
             "Nº palavras", "Nº caract. Short Desc.", 
             "Nº caract./palavra (média)", "Nº caract./palavra (máximo)", 
             "Frequência")

addWorkSheet -workbook $workbook -sheetName '16' -title $title  -headers $headers -inputData $stats["Desc. BD contém caracteres 'especiais'"] # -freq $stats["Frequencias"]["BD short desc."]

$time = StartStop-Print "2" $step $input $time;



############## sheet "Schemas sem short Desc nem Long Desc"


$step++ 
$ws_nb++
$title = "Schemas sem short Desc. nem Long Desc. preenchidas"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Short Desc.", "Long Desc.")

addWorkSheet -workbook $workbook -sheetName '21' -title $title -headers $headers -inputData $stats["Schemas sem short Desc nem Long Desc"]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "BD sem short Desc nem Long Desc"

$step++ 
$ws_nb++
$title = "Schemas com Long Desc. preenchida mas sem short Desc."
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;


$headers = @("BD", "Short Desc.", "Long Desc.")

addWorkSheet -workbook $workbook -sheetName '22' -title $title -headers $headers -inputData $stats["Schemas com Long Desc mas sem short Desc"]

$time = StartStop-Print "2" $step $input $time;



############## sheet "Schemas com short Desc"


$step++ 
$ws_nb++
$title = "Schemas com short Desc. preenchida"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Short Desc.", "Long Desc.",
             "Nº palavras", "Nº caract. Short Desc.", 
             "Nº caract./palavra (média)", "Nº caract./palavra (máximo)", 
             "Frequência")

addWorkSheet -workbook $workbook -sheetName '23' -title $title -headers $headers -inputData $stats["Schemas com short Desc"] # -freq $stats["Frequencias"]["Schema short desc."]

$time = StartStop-Print "2" $step $input $time;

############## sheet "Schemas com short Desc"


$step++ 
$ws_nb++
$title = "Tabelas com Short Desc. inválida (em nº caract./palavra)"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Short Desc.", "Long Desc.",            
             "Frequência")

addWorkSheet -workbook $workbook -sheetName '24' -title $title -headers $headers -inputData $stats["Desc. Schema inválida (em nº caract./palavra)"] # -freq $stats["Frequencias"]["Schema short desc."]

$time = StartStop-Print "2" $step $input $time;

############## sheet "Schemas com short Desc"


$step++ 
$ws_nb++
$title = "Tabelas com Short Desc. inválida (em nº palavras)"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Short Desc.", "Long Desc.",
             "Nº palavras", "Nº caract. Short Desc.", 
             "Nº caract./palavra (média)", "Nº caract./palavra (máximo)", 
             "Frequência")

addWorkSheet -workbook $workbook -sheetName '25' -title $title -headers $headers -inputData $stats["Desc. Schema inválida (em nº palavras)"] # -freq $stats["Frequencias"]["Schema short desc."]

$time = StartStop-Print "2" $step $input $time;

############## sheet "Schemas com short Desc"


$step++ 
$ws_nb++
$title = "Schemas - Short Desc. contém caracteres 'especiais'"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Short Desc.", "Long Desc.",
             "Nº palavras", "Nº caract. Short Desc.", 
             "Nº caract./palavra (média)", "Nº caract./palavra (máximo)", 
             "Frequência")

addWorkSheet -workbook $workbook -sheetName '26' -title $title -headers $headers -inputData $stats["Desc. Schema contém caracteres 'especiais'"] # -freq $stats["Frequencias"]["Schema short desc."]


$time = StartStop-Print "2" $step $input $time;

############## Sheet "Tabelas sem short Desc nem Long Desc"


$step++ 
$ws_nb++
$input = "Excel Sheet #"+$ws_nb+": 'Tabelas sem short Desc. nem Long Desc. preenchidas'"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Tabela", "Short Desc.", "Long Desc.")

addWorkSheet -workbook $workbook -sheetName '31' -title "Tabelas sem short Desc. nem Long Desc. preenchidas" -headers $headers -inputData $stats["Tabelas sem short Desc nem Long Desc"]

$time = StartStop-Print "2" $step $input $time;
  

############## Sheet "Tabelas com Long Desc mas sem short Desc"

$step++ 
$ws_nb++
$title = "Tabelas com Long Desc. preenchida mas sem short Desc."
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Tabela", "Short Desc.", "Long Desc.")

addWorkSheet -workbook $workbook -sheetName '32' -title $title -headers $headers -inputData $stats["Tabelas com Long Desc mas sem short Desc"]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "Tabelas com Short Desc"

$step++ 
$ws_nb++
$title = "Tabelas com Short Desc. preenchida"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Tabela", "Short Desc.", "Long Desc.", 
             "Nº palavras", "Nº caract. Short Desc.", "Nº caract./palavra (média)", "Nº caract./palavra (máximo)", "Frequência"
             )
addWorkSheet -workbook $workbook -sheetName '33' -title $title -headers $headers -inputData $stats["Tabelas com short Desc"] -freq $stats["Frequencias"]["Tab. short desc."]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "Tabelas com Short Desc. inválida (em nº caract./palavra)"

$step++ 
$ws_nb++
$title = "Tabelas com Short Desc. inválida (em nº caract./palavra)"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Tabela", "Short Desc.", "Long Desc.", 
             "Frequência"
             )
addWorkSheet -workbook $workbook -sheetName '34' -title $title -headers $headers -inputData $stats["Desc. Tabela inválida (em nº caract./palavra)"] -freq $stats["Frequencias"]["Tab. short desc."]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "Desc. Tabela inválida (em nº palavras)"

$step++ 
$ws_nb++
$title = "Tabelas com Short Desc. inválida (em nº palavras)"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Tabela", "Short Desc.", "Long Desc.", 
             "Nº palavras", "Nº caract. Short Desc.", "Nº caract./palavra (média)", "Nº caract./palavra (máximo)", "Frequência"
             )
addWorkSheet -workbook $workbook -sheetName '35' -title $title -headers $headers -inputData $stats["Desc. Tabela inválida (em nº palavras)"] -freq $stats["Frequencias"]["Tab. short desc."]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "Tabelas - Desc contém caracteres especiais

$step++ 
$ws_nb++
$title = "Tabelas - Short Desc. contém caracteres 'especiais'"
$input = "Excel Sheet #"+$ws_nb+": ""$title"""

$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Tabela", "Short Desc.", "Long Desc.")
addWorkSheet -workbook $workbook -sheetName '36' -title $title -headers $headers -inputData $stats["Desc. Tabela contém caracteres 'especiais'"]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "Colunas sem short Desc nem Long Desc"


$step++ 
$ws_nb++
$title = "Colunas sem short Desc. nem Long Desc. preenchidas"
$input = "Excel Sheet #"+$ws_nb+": '$title'"
$time = StartStop-Print "1" $step $input $time;


$headers = @("BD", "Schema", "Tabela", "Short Desc.", "Coluna")
addWorkSheet -workbook $workbook -sheetName '41' -title $title -headers $headers -inputData $stats["Colunas sem short Desc nem Long Desc"]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "Colunas com Long Desc mas sem short Desc"

$step++ 
$ws_nb++
$input = "Excel Sheet #"+$ws_nb+": ""Colunas com Long Desc mas sem short Desc"""
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Tabela", "Short Desc.",  "Coluna", "Short Desc.", "Long Desc.")

addWorkSheet -workbook $workbook -sheetName '42' -title "Colunas com Long Desc. preenchida mas sem short Desc." -headers $headers -inputData $stats["Colunas com Long Desc mas sem short Desc"]

$time = StartStop-Print "2" $step $input $time;



############## Sheet "Colunas com short Desc"


$step++ 
$ws_nb++
$title = "Colunas com short Desc. preenchida"
$input = "Excel Sheet #"+$ws_nb+": $title"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Tabela", "Short Desc.",  "Coluna", "Short Desc.", "Long Desc.", 
             "Nº palavras", "Nº caract. Short Desc.", "Nº caract./palavra (média)", "Nº caract./palavra (máximo)", "Frequência"
             )
addWorkSheet -workbook $workbook -sheetName '43' -title $title -headers $headers -inputData $stats["Colunas com short Desc"] -freq $stats["Frequencias"]["Col. short desc."]

$time = StartStop-Print "2" $step $input $time;



############## Sheet "Colunas - Desc too Short (in numb chars)"


$step++ 
$ws_nb++
$title = "Colunas com short Desc inválida em nº caract./palavra ou porque a descrição toma o valor (""Nome Coluna"", ""description"" )"
$input = "Excel Sheet #"+$ws_nb+": $title"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Tabela", "Short Desc." ,"Coluna", "Short Desc.", "Long Desc.", 
             "Nº palavras", "Nº caract. Short Desc.", "Nº caract./palavra (média)", "Nº caract./palavra (máximo)"
             "Frequência"
             )
addWorkSheet -workbook $workbook -sheetName '44' -title $title -headers $headers -inputData $stats["Desc. Coluna inválida (em nº caract./palavra)"] -freq $stats["Frequencias"]["Col. short desc."]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "Colunas - Desc too Short (in numb words)"


$step++ 
$ws_nb++
$title = "Colunas com short Desc inválida (em nº palavras)"
$input = "Excel Sheet #"+$ws_nb+": $title"
$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Tabela", "Short Desc." , "Coluna", "Short Desc.", "Long Desc.", 
             "Nº palavras", "Nº caract. Short Desc.", "Nº caract./palavra (média)", "Nº caract./palavra (máximo)", "Frequência"
             )
addWorkSheet -workbook $workbook -sheetName '45' -title $title -headers $headers -inputData $stats["Desc. Coluna inválida (em nº palavras)"] -freq $stats["Frequencias"]["Col. short desc."]

$time = StartStop-Print "2" $step $input $time;


############## Sheet "Colunas - Desc contém caracteres especiais

$step++ 
$ws_nb++
$input = "Excel Sheet #"+$ws_nb+": 'Colunas - Desc contém caracteres especiais'"

$time = StartStop-Print "1" $step $input $time;

$headers = @("BD", "Schema", "Tabela",  "Coluna", "Short Desc.", "Long Desc.")
addWorkSheet -workbook $workbook -sheetName '46' -title "Colunas - Short Desc. contém caracteres 'especiais'" -headers $headers -inputData $stats["Desc. Coluna contém caracteres 'especiais'"]

$time = StartStop-Print "2" $step $input $time;




############## Sheet 2: Statistics Summary


$ws_nb++; $input = "Excel Sheet #"+$ws_nb+": 'Stats'"
$step++ ; $time = StartStop-Print "1" $step $input $time;

####

$workSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $workbook.Worksheets.Item(1)) 

$workSheet.Name = 'Stats'

$workSheet.Hyperlinks.Add($workSheet.Cells.Item(1,1), "", "'Regras'!A1", "Ir para a sheet Regras", "Regras")| Out-Null
$workSheet.Range("A1:F1").MergeCells = $true


# settings used to group rows
$numb_rows_per_group = 7

#
$workSheet.Cells.Item(2,2) = "Estatísticas Globais (O denomidador das % é contextual)"
$workSheet.Range("B2:G2").MergeCells = $true #| Out-null  


# set headers and values columns and init rows
$col_labels=2
$row_headers=3
#
$row_init=$row_headers + 1
$col_tots = $col_labels + 1
$col_percent = $col_tots + 1
$col_ref = $col_percent + 2

# Write headers 

$workSheet.Cells.Item($row_headers,$col_labels) = "Métrica"
$workSheet.Cells.Item($row_headers,3) = "Total"
$workSheet.Cells.Item($row_headers,4) = "%"
# jump column E (which is formatted at the end of section)
$workSheet.Cells.Item($row_headers,6) = "Lista disponível na sheet:"

$row_headers_range = $row_headers.ToString() + ":" + $row_headers.ToString()
$workSheet.Cells.Range($row_headers_range).Font.Bold=$True
$workSheet.Cells.Range($row_headers_range).HorizontalAlignment = -4108
$workSheet.Cells.Range($row_headers_range).VerticalAlignment = -4108

# ==================================

# ===> Host section / Headers 

# ==================================

$workSheet.Cells.Item($row_init+0,$col_labels) = 'Host'
$workSheet.Cells.Item($row_init+1,$col_labels) = 'Host sem short Desc nem Long Desc'
$workSheet.Cells.Item($row_init+2,$col_labels) = 'Host com Long Desc mas sem short Desc'
$workSheet.Cells.Item($row_init+3,$col_labels) = 'Host com short Desc'
$workSheet.Cells.Item($row_init+4,$col_labels) = "Short Desc. inválida (em nº caract./palavra)"
$workSheet.Cells.Item($row_init+5,$col_labels) = "Short Desc. inválida (em nº palavras)"
$workSheet.Cells.Item($row_init+6,$col_labels) = "Short Desc. contém caracteres 'especiais'"

# set some borders
$row = $row_init; $range = "B"+$row +":F"+ $row
$workSheet.Cells.Range($range).Interior.ColorIndex = 27
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16 # BordersIndex = 9 is the bottom border - check https://docs.microsoft.com/en-us/office/vba/api/excel.xlbordersindex 
$row = $row_init + 3; $range = ""+$row +":"+ $row
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16


# ===>  DataBase section / Values 


$workSheet.Cells.Item($row_init+0,$col_tots) = $stats["Host"].Count
$workSheet.Cells.Item($row_init+1,$col_tots) = $stats["Host sem short Desc nem Long Desc"].Count
$workSheet.Cells.Item($row_init+2,$col_tots) = $stats["Host com Long Desc mas sem short Desc"].Count
$workSheet.Cells.Item($row_init+3,$col_tots) = $stats["Host com short Desc"].Count

$workSheet.Cells.Item($row_init+4,$col_tots) = $stats["Desc. Host inválida (em nº caract./palavra)"].Count
$workSheet.Cells.Item($row_init+5,$col_tots) = $stats["Desc. Host inválida (em nº palavras)"].Count
$workSheet.Cells.Item($row_init+6,$col_tots) = $stats["Desc. Host contém caracteres 'especiais'"].Count

# Percentages

$workSheet.Cells.Item($row_init+0,$col_percent) = 100
$workSheet.Cells.Item($row_init+1,$col_percent) = [math]::Round($stats["Host sem short Desc nem Long Desc"].Count / $stats["Host"].Count  * 100, 1)
$workSheet.Cells.Item($row_init+2,$col_percent) = [math]::Round($stats["Host com Long Desc mas sem short Desc"].Count / $stats["Host"].Count  * 100, 1)
$workSheet.Cells.Item($row_init+3,$col_percent) = [math]::Round($stats["Host com short Desc"].Count / $stats["Host"].Count  * 100, 1)

if ( $stats["Host com short Desc"].Count -gt 0){
    $workSheet.Cells.Item($row_init+4,$col_percent) = [math]::Round($stats["Desc. Host inválida (em nº caract./palavra)"].Count / $stats["Host com short Desc"].Count  * 100, 1)
    $workSheet.Cells.Item($row_init+5,$col_percent) = [math]::Round($stats["Desc. Host inválida (em nº palavras)"].Count / $stats["Host com short Desc"].Count  * 100, 1)
    $workSheet.Cells.Item($row_init+6,$col_percent) = [math]::Round($stats["Desc. Host contém caracteres 'especiais'"].Count / $stats["Host com short Desc"].Count  * 100, 1)
}

# Set References for the Sheet with the comprehensive list (ilf there is data)
if ( $stats["Host sem short Desc nem Long Desc"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+1,$col_ref), "", "'Sheet 01'!A1", "Ir para a Sheet 01", "Sheet 01")| Out-Null
}
if ( $stats["Host com Long Desc mas sem short Desc"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+2,$col_ref), "", "'Sheet 02'!A1", "Ir para a Sheet 02", "Sheet 02")| Out-Null
}
if ( $stats["Host com short Desc"].Count -gt 0 -or $fullReport -eq 'Y') {
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+3,$col_ref), "", "'Sheet 03'!A1", "Ir para a Sheet 03", "Sheet 03")| Out-Null
}
if ( $stats["Desc. Host inválida (em nº caract./palavra)"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+4,$col_ref), "", "'Sheet 04'!A1", "Ir para a Sheet 04", "Sheet 04")| Out-Null
}
if ( $stats["Desc. Host inválida (em nº palavras)"].Count -gt 0 -or $fullReport -eq 'Y') {
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+5,$col_ref), "", "'Sheet 05'!A1", "Ir para a Sheet 05", "Sheet 05")| Out-Null
}
if ( $stats["Desc. Host contém caracteres 'especiais'"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+6,$col_ref), "", "'Sheet 06'!A1", "Ir para a Sheet 06", "Sheet 06")| Out-Null
}



# write a vertical Merged Label
$row_last = $row_init + $numb_rows_per_group - 1

$range = "A"+$row_init+":"+"A"+$row_last
$MergeCells = $workSheet.Range($range) 

$MergeCells.MergeCells = $true
$workSheet.Cells.Item($row_init,1) = "Host"
$workSheet.Cells.Range($range).Font.Bold=$True
$workSheet.Cells.Range($range).HorizontalAlignment = -4108
$workSheet.Cells.Range($range).VerticalAlignment = -4108
$workSheet.Cells.Range($range).Orientation = -4171



# ==================================

# ===> DataBase section / Headers 

# ==================================

$row_init=$row_init + $numb_rows_per_group


$workSheet.Cells.Item($row_init+0,$col_labels) = 'BD'
$workSheet.Cells.Item($row_init+1,$col_labels) = 'BD sem short Desc nem Long Desc'
$workSheet.Cells.Item($row_init+2,$col_labels) = 'BD com Long Desc mas sem short Desc'
$workSheet.Cells.Item($row_init+3,$col_labels) = 'BD com short Desc'
$workSheet.Cells.Item($row_init+4,$col_labels) = "Short Desc. inválida (em nº caract./palavra)"
$workSheet.Cells.Item($row_init+5,$col_labels) = 'Short Desc. inválida (em nº palavras)'
$workSheet.Cells.Item($row_init+6,$col_labels) = "Short Desc. contém caracteres 'especiais'"

# set some borders
# BordersIndex = 9 is the bottom border - check https://docs.microsoft.com/en-us/office/vba/api/excel.xlbordersindex 
$row = $row_init; $range = "B"+$row +":F"+ $row
$workSheet.Cells.Range($range).Interior.ColorIndex = 27
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16
$row = $row_init + 3; $range = ""+$row +":"+ $row
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16


# ===>  DataBase section / Values 


$workSheet.Cells.Item($row_init+0,$col_tots) = $stats["BD"].Count
$workSheet.Cells.Item($row_init+1,$col_tots) = $stats["BD sem short Desc nem Long Desc"].Count
$workSheet.Cells.Item($row_init+2,$col_tots) = $stats["BD com Long Desc mas sem short Desc"].Count
$workSheet.Cells.Item($row_init+3,$col_tots) = $stats["BD com short Desc"].Count

$workSheet.Cells.Item($row_init+4,$col_tots) = $stats["Desc. BD inválida (em nº caract./palavra)"].Count
$workSheet.Cells.Item($row_init+5,$col_tots) = $stats["Desc. BD inválida (em nº palavras)"].Count
$workSheet.Cells.Item($row_init+6,$col_tots) = $stats["Desc. BD contém caracteres 'especiais'"].Count

# Percentages

$workSheet.Cells.Item($row_init+0,$col_percent) = 100
$workSheet.Cells.Item($row_init+1,$col_percent) = [math]::Round($stats["BD sem short Desc nem Long Desc"].Count / $stats["BD"].Count  * 100, 1)
$workSheet.Cells.Item($row_init+2,$col_percent) = [math]::Round($stats["BD com Long Desc mas sem short Desc"].Count / $stats["BD"].Count  * 100, 1)
$workSheet.Cells.Item($row_init+3,$col_percent) = [math]::Round($stats["BD com short Desc"].Count / $stats["BD"].Count  * 100, 1)

if ( $stats["BD com short Desc"].Count -gt 0){
    $workSheet.Cells.Item($row_init+4,$col_percent) = [math]::Round($stats["Desc. BD inválida (em nº caract./palavra)"].Count / $stats["BD com short Desc"].Count  * 100, 1)
    $workSheet.Cells.Item($row_init+5,$col_percent) = [math]::Round($stats["Desc. BD inválida (em nº palavras)"].Count / $stats["BD com short Desc"].Count  * 100, 1)
    $workSheet.Cells.Item($row_init+6,$col_percent) = [math]::Round($stats["Desc. BD contém caracteres 'especiais'"].Count / $stats["BD com short Desc"].Count  * 100, 1)
}
# Set References for the Sheet with the comprehensive list (ilf there is data)
if ( $stats["BD sem short Desc nem Long Desc"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+1,$col_ref), "", "'Sheet 11'!A1", "Ir para a Sheet 11", "Sheet 11")| Out-Null
}
if ( $stats["BD com Long Desc mas sem short Desc"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+2,$col_ref), "", "'Sheet 12'!A1", "Ir para a Sheet 12", "Sheet 12")| Out-Null
}
if ( $stats["BD com short Desc"].Count -gt 0 -or $fullReport -eq 'Y') {
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+3,$col_ref), "", "'Sheet 13'!A1", "Ir para a Sheet 13", "Sheet 13")| Out-Null
}
if ( $stats["Desc. BD inválida (em nº caract./palavra)"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+4,$col_ref), "", "'Sheet 14'!A1", "Ir para a Sheet 14", "Sheet 14")| Out-Null
}
if ( $stats["Desc. BD inválida (em nº palavras)"].Count -gt 0 -or $fullReport -eq 'Y') {
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+5,$col_ref), "", "'Sheet 15'!A1", "Ir para a Sheet 15", "Sheet 15")| Out-Null
}
if ( $stats["Desc. BD contém caracteres 'especiais'"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+6,$col_ref), "", "'Sheet 16'!A1", "Ir para a Sheet 16", "Sheet 16")| Out-Null
}



# write a vertical Merged Label
$row_last = $row_init + $numb_rows_per_group - 1

$range = "A"+$row_init+":"+"A"+$row_last
$MergeCells = $workSheet.Range($range) 

$MergeCells.MergeCells = $true
$workSheet.Cells.Item($row_init,1) = "BD"
$workSheet.Cells.Range($range).Font.Bold=$True
$workSheet.Cells.Range($range).HorizontalAlignment = -4108
$workSheet.Cells.Range($range).VerticalAlignment = -4108
$workSheet.Cells.Range($range).Orientation = -4171

# ==================================

# Schemas section  / Headers

# ==================================


$row_init=$row_init + $numb_rows_per_group

$workSheet.Cells.Item($row_init+0,$col_labels) = 'Schemas'
$workSheet.Cells.Item($row_init+1,$col_labels) = 'Schemas sem short Desc nem Long Desc'
$workSheet.Cells.Item($row_init+2,$col_labels) = 'Schemas com Long Desc mas sem short Desc'
$workSheet.Cells.Item($row_init+3,$col_labels) = 'Schemas com short Desc'
$workSheet.Cells.Item($row_init+4,$col_labels) = "Short Desc. inválida (em nº caract./palavra)"
$workSheet.Cells.Item($row_init+5,$col_labels) = 'Short Desc. inválida (em nº palavras)'
$workSheet.Cells.Item($row_init+6,$col_labels) = "Short Desc. contém caracteres 'especiais'"


# set some borders
# BordersIndex = 9 is the bottom border - check https://docs.microsoft.com/en-us/office/vba/api/excel.xlbordersindex 

$row = $row_init; $range = "B"+$row +":F"+ $row
$workSheet.Cells.Range($range).Interior.ColorIndex = 27
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16
$row = $row_init + 3; $range = ""+$row +":"+ $row
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16


# ===>  Values (Schemas Section)

$workSheet.Cells.Item($row_init+0,$col_tots) = $stats["Schemas"].Count
$workSheet.Cells.Item($row_init+1,$col_tots) = $stats["Schemas sem short Desc nem Long Desc"].Count
$workSheet.Cells.Item($row_init+2,$col_tots) = $stats["Schemas com Long Desc mas sem short Desc"].Count
$workSheet.Cells.Item($row_init+3,$col_tots) = $stats["Schemas com short Desc"].Count

$workSheet.Cells.Item($row_init+4,$col_tots) = $stats["Desc. Schema inválida (em nº caract./palavra)"].Count
$workSheet.Cells.Item($row_init+5,$col_tots) = $stats["Desc. Schema inválida (em nº palavras)"].Count
$workSheet.Cells.Item($row_init+6,$col_tots) = $stats["Desc. Schema contém caracteres 'especiais'"].Count

# Percentages


$workSheet.Cells.Item($row_init+0,$col_percent) = 100
$workSheet.Cells.Item($row_init+1,$col_percent) = [math]::Round($stats["Schemas sem short Desc nem Long Desc"].Count / $stats["Schemas"].Count  * 100, 1)
$workSheet.Cells.Item($row_init+2,$col_percent) = [math]::Round($stats["Schemas com Long Desc mas sem short Desc"].Count / $stats["Schemas"].Count  * 100, 1)
$workSheet.Cells.Item($row_init+3,$col_percent) = [math]::Round($stats["Schemas com short Desc"].Count / $stats["Schemas"].Count  * 100, 1)
if ( $stats["Schemas com short Desc"].Count -gt 0){
    $workSheet.Cells.Item($row_init+4,$col_percent) = [math]::Round($stats["Desc. Schema inválida (em nº caract./palavra)"].Count / $stats["Schemas com short Desc"].Count  * 100, 1)
    $workSheet.Cells.Item($row_init+5,$col_percent) = [math]::Round($stats["Desc. Schema inválida (em nº palavras)"].Count / $stats["Schemas com short Desc"].Count  * 100, 1)
    $workSheet.Cells.Item($row_init+6,$col_percent) = [math]::Round($stats["Desc. Schema contém caracteres 'especiais'"].Count / $stats["Schemas com short Desc"].Count  * 100, 1)
}

# Set References for the Sheet with the comprehensive list (ilf there is data)
if ( $stats["Schemas sem short Desc nem Long Desc"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+1,$col_ref), "", "'Sheet 21'!A1", "Ir para a Sheet 21", "Sheet 21")| Out-Null
}
if ( $stats["Schemas com Long Desc mas sem short Desc"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+2,$col_ref), "", "'Sheet 22'!A1", "Ir para a Sheet 22", "Sheet 22")| Out-Null
}
if ( $stats["Schemas com short Desc"].Count -gt 0 -or $fullReport -eq 'Y') {
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+3,$col_ref), "", "'Sheet 23'!A1", "Ir para a Sheet 23", "Sheet 23")| Out-Null
}
if ( $stats["Desc. Schema inválida (em nº caract./palavra)"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+4,$col_ref), "", "'Sheet 24'!A1", "Ir para a Sheet 24", "Sheet 24")| Out-Null
}
if ( $stats["Desc. Schema inválida (em nº palavras)"].Count -gt 0 -or $fullReport -eq 'Y') {
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+5,$col_ref), "", "'Sheet 25'!A1", "Ir para a Sheet 25", "Sheet 25")| Out-Null
}
if ( $stats["Desc. Schema contém caracteres 'especiais'"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+6,$col_ref), "", "'Sheet 26'!A1", "Ir para a Sheet 26", "Sheet 26")| Out-Null
}


# write a vertical Merged Label
$row_last = $row_init + $numb_rows_per_group - 1

$range = "A"+$row_init+":"+"A"+$row_last
$MergeCells = $workSheet.Range($range) 

$MergeCells.MergeCells = $true
$workSheet.Cells.Item($row_init,1) = "Schemas"
$workSheet.Cells.Range($range).Font.Bold=$True
$workSheet.Cells.Range($range).HorizontalAlignment = -4108
$workSheet.Cells.Range($range).VerticalAlignment = -4108
$workSheet.Cells.Range($range).Orientation = -4171

# ==================================

# Tables section / Headers   

# ==================================


$row_init=$row_init + $numb_rows_per_group

$workSheet.Cells.Item($row_init+0,$col_labels) = 'Tabelas'
$workSheet.Cells.Item($row_init+1,$col_labels) = 'Tabelas sem short Desc nem Long Desc'
$workSheet.Cells.Item($row_init+2,$col_labels) = 'Tabelas com Long Desc mas sem short Desc'
$workSheet.Cells.Item($row_init+3,$col_labels) = 'Tabelas com short Desc'
$workSheet.Cells.Item($row_init+4,$col_labels) = "Short Desc. inválida (em nº caract./palavra)"
$workSheet.Cells.Item($row_init+5,$col_labels) = 'Short Desc. inválida (em nº palavras)'
$workSheet.Cells.Item($row_init+6,$col_labels) = "Short Desc. contém caracteres 'especiais'"

# set some borders (and first row in color)
# BordersIndex = 9 is the bottom border - check https://docs.microsoft.com/en-us/office/vba/api/excel.xlbordersindex 
$row = $row_init; $range = "B"+$row +":F"+ $row
$workSheet.Cells.Range($range).Interior.ColorIndex = 27
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16
$row = $row_init + 3; $range = ""+$row +":"+ $row
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16


# ===>  Values (Tables Section)


$workSheet.Cells.Item($row_init+0,$col_tots) = $stats["Tabelas"].Count
$workSheet.Cells.Item($row_init+1,$col_tots) = $stats["Tabelas sem short Desc nem Long Desc"].Count
$workSheet.Cells.Item($row_init+2,$col_tots) = $stats["Tabelas com Long Desc mas sem short Desc"].Count
$workSheet.Cells.Item($row_init+3,$col_tots) = $stats["Tabelas com short Desc"].Count

$workSheet.Cells.Item($row_init+4,$col_tots) = $stats["Desc. Tabela inválida (em nº caract./palavra)"].Count
$workSheet.Cells.Item($row_init+5,$col_tots) = $stats["Desc. Tabela inválida (em nº palavras)"].Count
$workSheet.Cells.Item($row_init+6,$col_tots) = $stats["Desc. Tabela contém caracteres 'especiais'"].Count

# Percentages


$workSheet.Cells.Item($row_init+0,$col_percent) = 100
$workSheet.Cells.Item($row_init+1,$col_percent) = [math]::Round($stats["Tabelas sem short Desc nem Long Desc"].Count / $stats["Tabelas"].Count  * 100, 1)
$workSheet.Cells.Item($row_init+2,$col_percent) = [math]::Round($stats["Tabelas com Long Desc mas sem short Desc"].Count / $stats["Tabelas"].Count  * 100, 1)
$workSheet.Cells.Item($row_init+3,$col_percent) = [math]::Round($stats["Tabelas com short Desc"].Count / $stats["Tabelas"].Count  * 100, 1)
if ( $stats["Tabelas com short Desc"].Count -gt 0){
    $workSheet.Cells.Item($row_init+4,$col_percent) = [math]::Round($stats["Desc. Tabela inválida (em nº caract./palavra)"].Count / $stats["Tabelas com short Desc"].Count  * 100, 1)
    $workSheet.Cells.Item($row_init+5,$col_percent) = [math]::Round($stats["Desc. Tabela inválida (em nº palavras)"].Count / $stats["Tabelas com short Desc"].Count  * 100, 1)
    $workSheet.Cells.Item($row_init+6,$col_percent) = [math]::Round($stats["Desc. Tabela contém caracteres 'especiais'"].Count / $stats["Tabelas com short Desc"].Count  * 100, 1)
}

# Set References for the Sheet with the comprehensive list (ilf there is data)
if ( $stats["Tabelas sem short Desc nem Long Desc"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+1,$col_ref), "", "'Sheet 31'!A1", "Ir para a sheet 31", "Sheet 31")| Out-Null
}
if ( $stats["Tabelas com Long Desc mas sem short Desc"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+2,$col_ref), "", "'Sheet 32'!A1", "Ir para a sheet 32", "Sheet 32")| Out-Null
}
if ( $stats["Tabelas com short Desc"].Count -gt 0 -or $fullReport -eq 'Y') {
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+3,$col_ref), "", "'Sheet 33'!A1", "Ir para a sheet 33", "Sheet 33")| Out-Null
}
if ( $stats["Desc. Tabela inválida (em nº caract./palavra)"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+4,$col_ref), "", "'Sheet 34'!A1", "Ir para a sheet 34", "Sheet 34")| Out-Null
}
if ( $stats["Desc. Tabela inválida (em nº palavras)"].Count -gt 0 -or $fullReport -eq 'Y') {
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+5,$col_ref), "", "'Sheet 35'!A1", "Ir para a sheet 35", "Sheet 35")| Out-Null
}
if ( $stats["Desc. Tabela contém caracteres 'especiais'"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+6,$col_ref), "", "'Sheet 36'!A1", "Ir para a sheet 36", "Sheet 36")| Out-Null
}


# set some borders.
# BordersIndex = 9 is the bottom border - check https://docs.microsoft.com/en-us/office/vba/api/excel.xlbordersindex 

$row = $row_init; $range = ""+$row +":"+ $row
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16

$row = $row_init+3; $range = ""+$row +":"+ $row
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16




# write a vertical Merged Label
$row_last = $row_init + $numb_rows_per_group - 1

$range = "A"+$row_init+":"+"A"+$row_last
$MergeCells = $workSheet.Range($range) 

$MergeCells.MergeCells = $true
$workSheet.Cells.Item($row_init,1) = "Tabelas"
$workSheet.Cells.Range($range).Font.Bold=$True
$workSheet.Cells.Range($range).HorizontalAlignment = -4108
$workSheet.Cells.Range($range).VerticalAlignment = -4108
$workSheet.Cells.Range($range).Orientation = -4171

# ==================================

# Columns section / Headers  

# ==================================

$row_init = $row_init + $numb_rows_per_group

$workSheet.Cells.Item($row_init+0,$col_labels) = 'Colunas'
$workSheet.Cells.Item($row_init+1,$col_labels) = 'Colunas sem short Desc nem Long Desc'
$workSheet.Cells.Item($row_init+2,$col_labels) = 'Colunas com Long Desc mas sem short Desc'
$workSheet.Cells.Item($row_init+3,$col_labels) = 'Colunas com short Desc'
$workSheet.Cells.Item($row_init+4,$col_labels) = "Short Desc. inválida (em nº caract./palavra)"
$workSheet.Cells.Item($row_init+5,$col_labels) = 'Short Desc. inválida (em nº palavras)'
$workSheet.Cells.Item($row_init+6,$col_labels) = "Short Desc. contém caracteres 'especiais'"

# set some borders and first row in color
# BordersIndex = 9 is the bottom border - check https://docs.microsoft.com/en-us/office/vba/api/excel.xlbordersindex 

$row = $row_init; $range = "B"+$row +":F"+ $row
$workSheet.Cells.Range($range).Interior.ColorIndex = 27
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16
$row = $row_init + 3; $range = ""+$row +":"+ $row
$workSheet.Cells.Range($range).Borders(9).ColorIndex = 16



# ===>  Values (Columns Section)

$workSheet.Cells.Item($row_init+0,$col_tots) = $stats["Colunas"].Count
$workSheet.Cells.Item($row_init+1,$col_tots) = $stats["Colunas sem short Desc nem Long Desc"].Count
$workSheet.Cells.Item($row_init+2,$col_tots) = $stats["Colunas com Long Desc mas sem short Desc"].Count
$workSheet.Cells.Item($row_init+3,$col_tots) = $stats["Colunas com short Desc"].Count

$workSheet.Cells.Item($row_init+4,$col_tots) = $stats["Desc. Coluna inválida (em nº caract./palavra)"].Count
$workSheet.Cells.Item($row_init+5,$col_tots) = $stats["Desc. Coluna inválida (em nº palavras)"].Count
$workSheet.Cells.Item($row_init+6,$col_tots) = $stats["Desc. Coluna contém caracteres 'especiais'"].Count


# Percentages


$workSheet.Cells.Item($row_init+0,$col_percent) = 100
$workSheet.Cells.Item($row_init+1,$col_percent) = [math]::Round($stats["Colunas sem short Desc nem Long Desc"].Count / $stats["Colunas"].Count  * 100, 1)
$workSheet.Cells.Item($row_init+2,$col_percent) = [math]::Round($stats["Colunas com Long Desc mas sem short Desc"].Count / $stats["Colunas"].Count  * 100, 1)
$workSheet.Cells.Item($row_init+3,$col_percent) = [math]::Round($stats["Colunas com short Desc"].Count / $stats["Colunas"].Count  * 100, 1)
if ( $stats["Colunas com short Desc"].Count -gt 0){
    $workSheet.Cells.Item($row_init+4,$col_percent) =  [math]::Round($stats["Desc. Coluna inválida (em nº caract./palavra)"].Count / $stats["Colunas com short Desc"].Count  * 100, 1)
    $workSheet.Cells.Item($row_init+5,$col_percent) =  [math]::Round($stats["Desc. Coluna inválida (em nº palavras)"].Count / $stats["Colunas com short Desc"].Count  * 100, 1)
    $workSheet.Cells.Item($row_init+6,$col_percent) =  [math]::Round($stats["Desc. Coluna contém caracteres 'especiais'"].Count / $stats["Colunas com short Desc"].Count  * 100, 1)
}
# Set References for the Sheet with the comprehensive list (ilf there is data)

if ( $stats["Colunas sem short Desc nem Long Desc"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+1,$col_ref), "", "'Sheet 41'!A1", "Ir para a sheet 41","Sheet 41")| Out-Null
}
if ( $stats["Colunas com Long Desc mas sem short Desc"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+2,$col_ref), "", "'Sheet 42'!A1", "Ir para a sheet 42","Sheet 42")| Out-Null
}
if ( $stats["Colunas com short Desc"].Count -gt 0 -or $fullReport -eq 'Y') {
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+3,$col_ref), "", "'Sheet 43'!A1", "Ir para a sheet 43","Sheet 43")| Out-Null
}
if ( $stats["Desc. Coluna inválida (em nº caract./palavra)"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+4,$col_ref), "", "'Sheet 44'!A1", "Ir para a sheet 44", "Sheet 44")| Out-Null
}
if ( $stats["Desc. Coluna inválida (em nº palavras)"].Count -gt 0 -or $fullReport -eq 'Y') {
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+5,$col_ref), "", "'Sheet 45'!A1", "Ir para a sheet 45", "Sheet 45")| Out-Null
}
if ( $stats["Desc. Coluna contém caracteres 'especiais'"].Count -gt 0 -or $fullReport -eq 'Y'){
    $workSheet.Hyperlinks.Add($workSheet.Cells.Item($row_init+6,$col_ref), "", "'Sheet 46'!A1", "Ir para a sheet 46", "Sheet 46")| Out-Null
}


# write a vertical Merged Label
$row_last = $row_init + $numb_rows_per_group - 1

$range = "A"+$row_init+":"+"A"+$row_last
$MergeCells = $workSheet.Range($range) 

$MergeCells.MergeCells = $true
$workSheet.Cells.Item($row_init,1) = "Colunas"
$workSheet.Cells.Range($range).Font.Bold=$True
$workSheet.Cells.Range($range).HorizontalAlignment = -4108   # center aligned
$workSheet.Cells.Range($range).VerticalAlignment   = -4108
$workSheet.Cells.Range($range).Orientation = -4171

# Adjust columns width
$usedRange = $workSheet.UsedRange()
$usedRange.EntireColumn.AutoFit() | Out-Null

# Set some custom settings
$workSheet.Cells.Range("F:F").HorizontalAlignment = -4152  # right aligned
$workSheet.Cells.Range("D:D").ColumnWidth = 8
$workSheet.Cells.Range("E:E").ColumnWidth = 2
# BordersIndex = 10 is the right border - check https://docs.microsoft.com/en-us/office/vba/api/excel.xlbordersindex 
# ColorIndex 16 is gray  # check http://dmcritchie.mvps.org/excel/colors.htm 
$workSheet.Cells.Range("E:E").Borders(10).ColorIndex = 16

$WorkSheet.Application.ActiveWindow.SplitColumn = 1
$WorkSheet.Application.ActiveWindow.SplitRow = 3
$WorkSheet.Application.ActiveWindow.FreezePanes = $true



$time = StartStop-Print "2" $step $input $time;

############## Sheet 1:  Rules Analysis Summary

$step++ 
$ws_nb++
$input = "Excel Sheet #"+$ws_nb+": 'Regras'"
$time = StartStop-Print "1" $step $input $time;

$workSheet= $workbook.Worksheets.Item(1)

# Freezing panes does not work on this sheet.... No fucking clue
# 
# $WorkSheet.Application.ActiveWindow.SplitColumn = 3
# $WorkSheet.Application.ActiveWindow.SplitRow = 3
# $WorkSheet.Application.ActiveWindow.FreezePanes = $true



$workSheet.Name = 'Regras'

$workSheet.Hyperlinks.Add($workSheet.Cells.Item(1,1), "", "'Stats'!A1", "Ir para a sheet Stats",    "Stats")| Out-Null

$row_headers=3
$row_init= $row_headers + 1

$col_labels = 2
$col_values = 3


$workSheet.Cells.Item(2,$col_labels) = "Regras sobre o campo 'Short Desc'"
$workSheet.Cells.Item($row_init+0,$col_labels)  = "Nº mínimo caracteres / palavra - Host"
$workSheet.Cells.Item($row_init+1,$col_labels)  = "Nº mínimo caracteres / palavra - BD"
$workSheet.Cells.Item($row_init+2,$col_labels)  = "Nº mínimo caracteres / palavra - Schemas"
$workSheet.Cells.Item($row_init+3,$col_labels)  = "Nº mínimo caracteres / palavra - Tabelas"
$workSheet.Cells.Item($row_init+4,$col_labels)  = "Nº mínimo caracteres / palavra - Colunas"

$workSheet.Cells.Item($row_init+0,$col_values)  = $MIN_HOST_DESC_NCHARS_PER_WORD
$workSheet.Cells.Item($row_init+1,$col_values)  = $MIN_DB_DESC_NCHARS_PER_WORD
$workSheet.Cells.Item($row_init+2,$col_values)  = $MIN_SCHE_DESC_NCHARS_PER_WORD
$workSheet.Cells.Item($row_init+3,$col_values)  = $MIN_TAB_DESC_NCHARS_PER_WORD
$workSheet.Cells.Item($row_init+4,$col_values)  = $MIN_COL_DESC_NCHARS_PER_WORD

$workSheet.Cells.Item($row_init+6,$col_labels)  = "Nº mínimo palavras - Host"
$workSheet.Cells.Item($row_init+7,$col_labels)  = "Nº mínimo palavras - BD"
$workSheet.Cells.Item($row_init+8,$col_labels)  = "Nº mínimo palavras - Schemas"
$workSheet.Cells.Item($row_init+9,$col_labels)  = "Nº mínimo palavras - Tabelas"
$workSheet.Cells.Item($row_init+10,$col_labels)  = "Nº mínimo palavras - Colunas"

$workSheet.Cells.Item($row_init+6,$col_values)  = $MIN_HOST_DESC_NWORDS
$workSheet.Cells.Item($row_init+7,$col_values)  = $MIN_DB_DESC_NWORDS
$workSheet.Cells.Item($row_init+8,$col_values)  = $MIN_SCHE_DESC_NWORDS
$workSheet.Cells.Item($row_init+9,$col_values)  = $MIN_TAB_DESC_NWORDS
$workSheet.Cells.Item($row_init+10,$col_values)  = $MIN_COL_DESC_NWORDS

$workSheet.Cells.Range("B4:B14").Font.Bold=$True


$workSheet.Cells.Item(2,5) = "Lista Palavras permitidas para o campo 'Short Desc' de 1 palavra *`r`n*Excepções ao critério do nº mínimo de palavras"
$workSheet.Range("E2:I2").MergeCells = $true #| Out-null

$row_headers_range = $row_headers.ToString() + ":" + $row_headers.ToString()
$workSheet.Cells.Range($row_headers_range).Font.Bold=$True
$workSheet.Cells.Range($row_headers_range).HorizontalAlignment = -4108
$workSheet.Cells.Range($row_headers_range).VerticalAlignment = -4108

$workSheet.Cells.Item($row_headers,5) = "Host"
$workSheet.Cells.Item($row_headers,6) = "BD"
$workSheet.Cells.Item($row_headers,7) = "Schemas"
$workSheet.Cells.Item($row_headers,8) = "Tabelas"
$workSheet.Cells.Item($row_headers,9) = "Colunas"

#######

$workSheet.Cells.Item(2,10) = "Lista Palavras não permitidas sobre o campo 'Short Desc' de 1 palavra *`r`n*Regra aplicada mesmo no caso de admitirmos descrições de 1 palavra"
$workSheet.Range("J2:N2").MergeCells = $true #| Out-null
$workSheet.Rows("2").RowHeight = 30


$row_headers_range = $row_headers.ToString() + ":" + $row_headers.ToString()
$workSheet.Cells.Range($row_headers_range).Font.Bold=$True
$workSheet.Cells.Range($row_headers_range).HorizontalAlignment = -4108
$workSheet.Cells.Range($row_headers_range).VerticalAlignment = -4108

$workSheet.Cells.Item($row_headers,10) = "Host"
$workSheet.Cells.Item($row_headers,11) = "BD"
$workSheet.Cells.Item($row_headers,12) = "Schemas"
$workSheet.Cells.Item($row_headers,13) = "Tabelas"
$workSheet.Cells.Item($row_headers,14) = "Colunas"

# loop on the lists of words

# Allowed
$row = $row_headers + 1
foreach ($pattern in $host_short_desc_allowed){    
    $workSheet.Cells.Item($row,5) = $pattern
    $row++
}
$row = $row_headers + 1
foreach ($pattern in $db_short_desc_allowed){    
    $workSheet.Cells.Item($row,6) = $pattern
    $row++
}
$row = $row_headers + 1
foreach ($pattern in $schema_short_desc_allowed){    
    $workSheet.Cells.Item($row,7) = $pattern
    $row++
}
$row = $row_headers + 1
foreach ($pattern in $tab_short_desc_allowed){    
    $workSheet.Cells.Item($row,8) = $pattern
    $row++
}
$row = $row_headers + 1
foreach ($pattern in $col_short_desc_allowed){    
    $workSheet.Cells.Item($row,9) = $pattern
    $row++
}

# Forbidden
$row = $row_headers + 1
foreach ($pattern in $host_short_desc_forbidden){    
    $workSheet.Cells.Item($row,10) = $pattern
    $row++
}
$row = $row_headers + 1
foreach ($pattern in $db_short_desc_forbidden){    
    $workSheet.Cells.Item($row,11) = $pattern
    $row++
}
$row = $row_headers + 1
foreach ($pattern in $schema_short_desc_forbidden){    
    $workSheet.Cells.Item($row,12) = $pattern
    $row++
}
$row = $row_headers + 1
foreach ($pattern in $tab_short_desc_forbidden){    
    $workSheet.Cells.Item($row,13) = $pattern
    $row++
}
$row = $row_headers + 1
foreach ($pattern in $col_short_desc_forbidden){    
    $workSheet.Cells.Item($row,14) = $pattern
    $row++
}


# set Vertical Borders (on the right side of the selected column)
# ---------------------------------------------------------------
# BordersIndex = 7 is the left border - check https://docs.microsoft.com/en-us/office/vba/api/excel.xlbordersindex 
# ColorIndex 16 is gray  # check http://dmcritchie.mvps.org/excel/colors.htm 

$workSheet.Cells.Range("C:C").Borders(10).ColorIndex = 16 # 16 is gray  # check http://dmcritchie.mvps.org/excel/colors.htm 
$workSheet.Cells.Range("D:D").Borders(10).ColorIndex = 16 # 16 is gray  # check http://dmcritchie.mvps.org/excel/colors.htm 
$workSheet.Cells.Range("I:I").Borders(10).ColorIndex = 16 # 16 is gray  # check http://dmcritchie.mvps.org/excel/colors.htm 
$workSheet.Cells.Range("N:N").Borders(10).ColorIndex = 16 # 16 is gray  # check http://dmcritchie.mvps.org/excel/colors.htm 

# Adjust automatically columns width
$usedRange = $workSheet.UsedRange()
$usedRange.EntireColumn.AutoFit() | Out-Null

# override the automatic auto-adjustment width
$workSheet.Cells.Range("E:N").ColumnWidth = 15

# set outside borders for region "D2:K2"
$workSheet.Cells.Range("E2:N2").Borders(7).ColorIndex = 16
$workSheet.Cells.Range("E2:N2").Borders(8).ColorIndex = 16
$workSheet.Cells.Range("E2:N2").Borders(9).ColorIndex = 16
$workSheet.Cells.Range("E2:N2").Borders(10).ColorIndex = 16
# set inner vertical border for range "D2:K2"
$workSheet.Cells.Range("J2").Borders(7).ColorIndex = 16
# set center alignment
$workSheet.Cells.Range("E2:N2").HorizontalAlignment = -4108



$time = StartStop-Print "2" $step $input $time;


#################   ENDING CONTROLS    ############## 


$time = Get-Date
$dt_tot = $time - $time0

$dt_min = $dt_tot.Minutes
$dt_sec = $dt_tot.Seconds


Write-Host "Writing output to Excel file $XLS_outfile at $time ..." -ForegroundColor Green
Write-Host "Total execution time: $dt_min min $dt_sec sec" -ForegroundColor Green
#saving & closing the file (activating second sheet)
$workbook.Worksheets.Item(2).Activate()
# Select cell A2
$workbook.Worksheets.Item(2).Cells.Range("A2").Select() | Out-null

$excel.ActiveWorkbook.SaveAs($XLS_outfile)
$excel.ActiveWorkbook.Close()
$excel.Quit()
