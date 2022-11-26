#!/usr/bin/perl

use strict;
use warnings;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;
use Date::Formatter;

my $row=0;
my $k;
my $i;
my $m;
my $b;
my @T_TO;
my @T_LDG;
my @Time_TO;
my @Time_LDG;
my @Out;
my $out_file;
my @FlightNum;
my @Flights;
my @Title_1;
my @Reg;
my @Regs;
my @elements;
my @strings;
my @Day;
my @Mo;
my @Ye;
my @Day_TO;
my @Mo_TO;
my @Ye_TO;
my @N2L;
my @N2R;


my @S001;
my @S002;
my @ENGS_ST_RUN_TIME;#массив для записи времени с начала запуска до N2<55% (для обоих двигателей)
my @TIME_ENG_START;
my @TIME_ENG_STOP;
my @CHECK;
my @TYPE;
my @AC_TYPE;
							my $CAS_everage =0;
							my $i_cas;
							my $GS_everage = 0;
							my $i_fuel;
							my $run;
							my $run_end;
							my $z;
							my @FF1;
							my @FF2;
							my @CWt;
							my @RUN_TIME;
							my @TIME_RUN;
							my @RUN_TIME_END;
							my @RUN_END_TIME;
							my @CAS_TO_TABLE;
							my @GS_TO_TABLE;
							my @CAS;
							my @FUEL_1;
							my @FUEL_2;
							my @FUEL_3;
							my @FUEL_4;
							my @FUEL_5;
							my $Fuel_cons_1;#расход топлива до руления перед взлетом.
							my $Fuel_cons_2;#расход топлива на рулении перед взлетом
							my $Fuel_cons_3;#расход топлива в полете
							my $Fuel_cons_4;#расход топлива на пробеге
							my $Fuel_cons_5;#расход топлива на рулении после посадки
					
							my $Fuel_cons_1_f;
                            my $Fuel_cons_2_f;
							my $Fuel_cons_3_f;
							my $Fuel_cons_4_f;
							my $Fuel_cons_5_f;
							my $ATOW;
							my @ATOW_TO_TABLE;
							my @Route;
																	
my $day_f;
my $mo_f;
my $ye_f;
my $dlt;


my @nms = glob"*.txt";

for($k = 0;$k <= $#nms;$k++)
{
                                                     @T_TO = 0;
													 @T_LDG = 0;
													 @Day = "";
                                                     @Mo = "";
                                                     @Ye = "";
													 @N2L = "";
													 @N2R = "";
													
													 
													 @S001 = "";
													 @S002 = "";
													 @ENGS_ST_RUN_TIME = 0;
													 @RUN_TIME = 0;#!!!!!!
													
												
													 $i_cas = 0;
													 $i_fuel = 0;
													 $Fuel_cons_1 = 0;
													 $Fuel_cons_2 = 0;
													 $Fuel_cons_3 = 0;
													 $Fuel_cons_4 = 0;
													 $Fuel_cons_5 = 0;
													 $run = 0;
													 $run_end = 0;
													 $m = 0;
													 
													 
	@Out = split(/\./,$nms[$k]);
	$Out[0] =~ s/\-/\_/g;
    $Out[0] =~ s/ /\_/g;
	$Out[0] =~ s/ABG//g;
	$Out[0] =~ s/\(//g;
	$Out[0] =~ s/\)//g;
	$out_file = $Out[0];
	@FlightNum = split(/_/,$out_file);
	push @Flights,"RL-".$FlightNum[0];
	push @Route,$FlightNum[1]."-".$FlightNum[2];
		
		open (my $fh,$nms[$k]) or die "Could not open file";
	
		
		 @strings = <$fh>;
		       @Title_1 = split(/\[/,$strings[0]);
               @Reg = split(/\]/,$Title_1[2]);
			   @TYPE = split(/\]/,$Title_1[1]);
	push @Regs,$Reg[0];
    push @AC_TYPE,$TYPE[0];	
	
	
		 for($m=30;$m <= $#strings;$m++)

        {
			if($AC_TYPE[$k] =~ /737/gi) {              #####   737-800
		
		
                                              chomp $strings[$m];
                                              @elements = split(/\t/,$strings[$m]);
											  
                        if($elements[21] == 1)
						                     { push @T_TO,$elements[0];
						                       push @Day,$elements[1];
											   push @Mo,$elements[2];
											   push @Ye,$elements[3];
						                     }
						
						if($elements[22] == 1){push @T_LDG,$elements[0];$run_end = $m}
						
						if($elements[13] == 1 or $elements[14] == 1){push @ENGS_ST_RUN_TIME,$elements[0];}
						
						if($elements[10] == 1){$i_cas++;$CAS_everage = $CAS_everage + $elements[5];$GS_everage = $GS_everage + $elements[6];}
						
						if($m == 53){$ATOW = $elements[4];$z = "check";}
						
						
						if($elements[16] == 1 and $m < 2000){push @RUN_TIME,$elements[0]}
						
						if($elements[22] == 0 and $elements[6] > 2 and $m > $run_end and $run_end != 0){push @RUN_TIME_END,$elements[0]}
						
						if($elements[19] == 0 and $elements[11] == 0){$elements[7] = 0}#защита от "пилы"
						if($elements[20] == 0 and $elements[12] == 0){$elements[8] = 0}
						
		if(($elements[11] == 1 or $elements[12] == 1) and $elements[17] != 1 and $elements[18] != 1 and $elements[9] != 1 and $elements[10] != 1 and $m < 1500){$Fuel_cons_1 = $Fuel_cons_1 + $elements[7] + $elements[8]}#check!!!!
		
		
		if($elements[17] == 1 or $elements[18] == 1){$Fuel_cons_2 = $Fuel_cons_2 + $elements[7] + $elements[8]}#check!!!!
		
		
		if($elements[10] == 1){$i_fuel++;$Fuel_cons_3 = $Fuel_cons_3 + $elements[7] + $elements[8]}
		
		
		if($elements[22] == 1){$run = $m;$Fuel_cons_4 = $Fuel_cons_4 + $elements[7] + $elements[8]}
		
		
		if($elements[10] == 0 and $m > $run and $run != 0){$Fuel_cons_5 = $Fuel_cons_5 + $elements[7] + $elements[8]}
						                                               
									}
									             ########  end 737-800                                                   
												                                                        #else { 
		                                                                                                       if($AC_TYPE[$k] =~ /757/gi) { #####757-200
									
									                                                                            chomp $strings[$m];
                                                                                                                @elements = split(/\t/,$strings[$m]);
																												
                                                                                                        if($elements[21] == 1)
																										      { push @T_TO,$elements[0];
						                                                                                        push @Day,$elements[1];
											                                                                    push @Mo,$elements[2];
											                                                                    push @Ye,$elements[3];
						                                                                                      }
						
						                                                                                if($elements[18] == 1){push @T_LDG,$elements[0];$run_end = $m;}
			                                                                                            if($elements[13] == 1 or $elements[14] == 1){push @ENGS_ST_RUN_TIME,$elements[0];}
										if($elements[12] == 1){$i_cas++;$CAS_everage = $CAS_everage + $elements[4];$GS_everage = $GS_everage + $elements[5];}
										
									    if($m == 50){$ATOW = $elements[8]}
										
										
										if($elements[14] == 1 and $elements[16] == 0 and $elements[12] == 0 and $m < 1500){$Fuel_cons_1 = $Fuel_cons_1 + $elements[6] + $elements[7]}	
										
										
										if($elements[16] == 1 and $m < 2000){$Fuel_cons_2 = $Fuel_cons_2 + $elements[6] + $elements[7]}
										if($elements[16] == 1 and $m < 2000){push @RUN_TIME,$elements[0]}
										
										if($elements[12] == 1){$Fuel_cons_3 = $Fuel_cons_3 + $elements[6] + $elements[7]}
										if($elements[18] == 1){$run = $m;$Fuel_cons_4 = $Fuel_cons_4 + $elements[6] + $elements[7]}
										if($elements[13] == 1 and $m > $run and $run != 0){$Fuel_cons_5 = $Fuel_cons_5 + $elements[6] + $elements[7]}
										
										if($elements[18] != 1 and $elements[5] > 2 and $m > $run_end and $run_end != 0){push @RUN_TIME_END,$elements[0]}
										
																										
			                                                                                                                             }
						                                                                                    													
														
														if($AC_TYPE[$k] =~ /767/gi) {   ##### 767-300
														chomp $strings[$m];
														@elements = split(/\t/,$strings[$m]);
														my $Ye_767 = 20;
														if($elements[18] == 1){push @T_TO,$elements[0];
														                      push @Day,$elements[1];
														                      push @Mo,$elements[2];
																			  push @Ye,$Ye_767;
														                     }
																			 if($elements[20] == 1){push @T_LDG,$elements[0];$run_end = $m}
														                     if($elements[13] == 1 or $elements[14] == 1){push @ENGS_ST_RUN_TIME,$elements[0];}
																			 if($elements[15] == 1){push @RUN_TIME,$elements[0]}
																			 if($elements[20] != 1 and $elements[5] > 2 and $m > $run_end and $run_end != 0){push @RUN_TIME_END,$elements[0]}
						 if($elements[12] == 1 and $elements[15] == 0 and $elements[16] == 0 and $elements[10] == 0 and $elements[11] == 0 and $m < 1500){$Fuel_cons_1 = $Fuel_cons_1 + $elements[6] + $elements[7]}
						 
						 if($elements[15] == 1 or $elements[16] == 1){$Fuel_cons_2 = $Fuel_cons_2 + $elements[6] + $elements[7]}
						 
						 if($elements[10] == 1 or $elements[11] == 1){$Fuel_cons_3 = $Fuel_cons_3 + $elements[6] + $elements[7]}
						 
						 if($elements[20] == 1){$run = $m;$Fuel_cons_4 = $Fuel_cons_4 + $elements[6] + $elements[7]}	
						
						 if($elements[12] == 1 and $m > $run and $run != 0){$Fuel_cons_5 = $Fuel_cons_5 + $elements[6] + $elements[7]}
						
						 if($elements[11] == 1){$i_cas++;$CAS_everage = $CAS_everage + $elements[4];$GS_everage = $GS_everage + $elements[5];}
						
						 if($m == 50){$ATOW = $elements[3]}

						
														                            }    ##### end 767-300
									  																		   													   
								}
                    
                   for($i = 900;$i >= 0;$i--)
					 
    {
	
	if((($ENGS_ST_RUN_TIME[$i] - $ENGS_ST_RUN_TIME[$i -1] > 3) and ($ENGS_ST_RUN_TIME[$i] - $ENGS_ST_RUN_TIME[$i -1] < 86397))
	  or (($ENGS_ST_RUN_TIME[$i] - $ENGS_ST_RUN_TIME[$i -1] < -15)
	  and ($ENGS_ST_RUN_TIME[$i] - $ENGS_ST_RUN_TIME[$i -1] > -86397)
	  )
	                                                         ){delete$ENGS_ST_RUN_TIME[$i-1];$ENGS_ST_RUN_TIME[$i] = $ENGS_ST_RUN_TIME[$i +1];
	                                                           $ENGS_ST_RUN_TIME[$i-1] = $ENGS_ST_RUN_TIME[$i]}
	 }
	
					for($b = ($#ENGS_ST_RUN_TIME - 900);$b < $#ENGS_ST_RUN_TIME;$b++)
	{
	if
	(
	(($ENGS_ST_RUN_TIME[$b + 1] - $ENGS_ST_RUN_TIME[$b]> 15) and ($ENGS_ST_RUN_TIME[$b + 1] - $ENGS_ST_RUN_TIME[$b] < 86396)) or
	(($ENGS_ST_RUN_TIME[$b + 1] - $ENGS_ST_RUN_TIME[$b] < 1) and ($ENGS_ST_RUN_TIME[$b + 1] - $ENGS_ST_RUN_TIME[$b] > -86396))
	)
	{delete$ENGS_ST_RUN_TIME[$b+1];$ENGS_ST_RUN_TIME[$b + 1] = $ENGS_ST_RUN_TIME[$b];$ENGS_ST_RUN_TIME[$b] = $ENGS_ST_RUN_TIME[$b - 1]}
	
	}

                             

	                                         if($Regs[$k] =~ /BRE/gi){$CAS_everage = $CAS_everage*0.539956803455724}else{$CAS_everage = $CAS_everage}
                                                 $CAS_everage = $CAS_everage/$i_cas;
												 $GS_everage = $GS_everage/$i_cas;
										         $Fuel_cons_1 = $Fuel_cons_1*0.45359237/3600;
										         $Fuel_cons_2 = $Fuel_cons_2*0.45359237/3600;
										         $Fuel_cons_3 = $Fuel_cons_3*0.45359237/3600;
										         $Fuel_cons_4 = $Fuel_cons_4*0.45359237/3600;
                                                 $Fuel_cons_5 = $Fuel_cons_5*0.45359237/3600;
												 
												if($Regs[$k] =~ /BLC/gi or $Regs[$k] =~ /BLG/gi) {$ATOW = $ATOW*0.45359237/1000;}else{$ATOW = $ATOW}
												 
		
		push @Time_TO,$T_TO[1];
		push @Day_TO,$Day[1];
		push @Mo_TO,$Mo[1];
		push @Ye_TO,$Ye[1];
		push @Time_LDG,$T_LDG[$#T_LDG];
		push @TIME_ENG_START,$ENGS_ST_RUN_TIME[1];
		push @TIME_ENG_STOP,$ENGS_ST_RUN_TIME[$#ENGS_ST_RUN_TIME];
		push @TIME_RUN,$RUN_TIME[1];
		push @RUN_END_TIME,$RUN_TIME_END[$#RUN_TIME_END-2];
		push @CAS_TO_TABLE,$CAS_everage;
		push @GS_TO_TABLE,$GS_everage;
		push @FUEL_1,$Fuel_cons_1;
		push @FUEL_2,$Fuel_cons_2;
		push @FUEL_3,$Fuel_cons_3;
		push @FUEL_4,$Fuel_cons_4;
		push @FUEL_5,$Fuel_cons_5;
		push @ATOW_TO_TABLE,$ATOW;
		
		close $fh;
		
}

 my $parser = Spreadsheet::ParseExcel->new(
        CellHandler => \&cell_handler,
        NotSetCell  => 1
    );

                    my $workbook = $parser->parse('START_STOP.xls');
                    sub cell_handler {
                    my $workbook    = $_[0];
                    my $sheet_index = $_[1];
                    $row = $_[2];
                    my $col = $_[3];
                    my $cell = $_[4];
                 	  
		   }
	   
	         
	my $count = $row; 	
		  
   $parser = Spreadsheet::ParseExcel::SaveParser->new();
my $template = $parser->Parse('START_STOP.xls');

my $worksheet = $template -> worksheet(0);
 $row = 0;
my $col = 0;
for($i = 0;$i<=$#Flights;$i++){     
										$Fuel_cons_1_f = sprintf("%0.2f",$FUEL_1[$i]);
										$Fuel_cons_2_f = sprintf("%0.2f",$FUEL_2[$i]);
										$Fuel_cons_3_f = sprintf("%0.2f",$FUEL_3[$i]);
										$Fuel_cons_4_f = sprintf("%0.2f",$FUEL_4[$i]);
										$Fuel_cons_5_f = sprintf("%0.2f",$FUEL_5[$i]);
										$CAS_TO_TABLE[$i] = sprintf("%0.2f",$CAS_TO_TABLE[$i]);
										$GS_TO_TABLE[$i] = sprintf("%0.2f",$GS_TO_TABLE[$i]);
										$ATOW_TO_TABLE[$i] = sprintf("%0.2f",$ATOW_TO_TABLE[$i]);
							   
                         $day_f = sprintf("%02d",$Day_TO[$i]);
						 $mo_f = sprintf("%02d",$Mo_TO[$i]);
						 $ye_f = sprintf("%02d",$Ye_TO[$i]);
						 
						 
                          my $hour = $Time_TO[$i]/3600;					 
                          my $hour_r = sprintf("%02d",$hour);
					      my $sec = $Time_TO[$i] - ($hour_r*3600);
						  my $min = $sec/60;
                          my $min_r = sprintf("%02d",$min);
						  
						  my $hour_ldg = $Time_LDG[$i]/3600;
						  my $hour_ldg_r = sprintf("%02d",$hour_ldg);
						  my $sec_ldg = $Time_LDG[$i] - ($hour_ldg_r*3600);
						  my $min_ldg = $sec_ldg/60;
						  my $min_ldg_r = sprintf("%02d",$min_ldg);
						  
						  my $DTime;
						  if ($Time_LDG[$i] > $Time_TO[$i]){$DTime = $Time_LDG[$i] - $Time_TO[$i]} else {$DTime = (86400 - $Time_TO[$i]) + $Time_LDG[$i]}
						  
						  my $Dhour = $DTime/3600;
						  my $Dhour_r = sprintf("%02d",$Dhour);
						  my $Dsec = $DTime - ($Dhour_r*3600);
						  my $Dmin = $Dsec/60;
						  my $Dmin_r = sprintf("%02d",$Dmin);
						  
						  
						                                  										
															
															my $hour_start = $TIME_ENG_START[$i]/3600;
															my $hour_start_r = sprintf("%02d",$hour_start);
															my $sec_start = $TIME_ENG_START[$i] - ($hour_start_r*3600);
															my $min_start = $sec_start/60;
															my $min_start_r = sprintf("%02d",$min_start);
															
															
															my $hour_stop = $TIME_ENG_STOP[$i]/3600;
															my $hour_stop_r = sprintf("%02d",$hour_stop);
															my $sec_stop = $TIME_ENG_STOP[$i] - ($hour_stop_r*3600);
															my $min_stop = $sec_stop/60;
															my $min_stop_r = sprintf("%02d",$min_stop);
															
															
															my $Start_and_Run_Time;
															
							if ($TIME_ENG_STOP[$i] > $TIME_ENG_START[$i]){$Start_and_Run_Time = $TIME_ENG_STOP[$i] - $TIME_ENG_START[$i]} else {$Start_and_Run_Time = (86400 - $TIME_ENG_START[$i]) + $TIME_ENG_STOP[$i]}
						  
						                                    my $Delta_hour = $Start_and_Run_Time/3600;
															my $Delta_hour_r = sprintf("%02d",$Delta_hour);
															my $Delta_sec = $Start_and_Run_Time - ($Delta_hour_r*3600);
															my $Delta_min = $Delta_sec/60;
															my $Delta_min_r = sprintf("%02d",$Delta_min);
															
															
															my $start_run_hour = $TIME_RUN[$i]/3600;
															my  $start_run_hour_r = sprintf("%02d",$start_run_hour);
															my $sec_start_run = $TIME_RUN[$i] - ($start_run_hour_r*3600);
															my $min_start_run = $sec_start_run/60;
															my $min_start_run_r = sprintf("%02d",$min_start_run);
															 
															 
															my $run_end_hour = $RUN_END_TIME[$i]/3600;
															my $run_end_hour_r = sprintf("%02d",$run_end_hour);
															my $sec_run_end = $RUN_END_TIME[$i] - ($run_end_hour_r*3600);
															my $min_run_end = $sec_run_end/60;
															my $min_start_end_r = sprintf("%02d",$min_run_end);
															
			if(((($hour_stop - $hour_ldg) > 1) and  ($hour_stop - $hour_ldg) < 23)or (($hour_stop - $hour_ldg) < -1) and  (($hour_stop - $hour_ldg) > -23)){$Delta_min_r = "";$Delta_hour_r = "check data"}
			if(((($hour_start - $hour) > 1) and  ($hour_start - $hour) < 23)or (($hour_start - $hour) < -1) and  (($hour_start - $hour) > -23)){$Delta_min_r = "";$Delta_hour_r = "check data"}
																				  

my $cell = $worksheet->get_cell($row+2,$col);
my $format_number = $cell->{FormatNo};
$worksheet->AddCell($row+$count+1+$i,$col+1,$Regs[$i],$format_number);
$worksheet->AddCell($row+$count+1+$i,$col+2,$Flights[$i],$format_number);
$worksheet->AddCell($row+$count+1+$i,$col+4,"$start_run_hour_r\:$min_start_run_r",$format_number);#run_start
$worksheet->AddCell($row+$count+1+$i,$col+5,"$hour_r\:$min_r",$format_number);#take off
$worksheet->AddCell($row+$count+1+$i,$col+6,"$hour_ldg_r\:$min_ldg_r",$format_number);#landing
$worksheet->AddCell($row+$count+1+$i,$col,"$day_f-$mo_f-20$ye_f",$format_number);
$worksheet->AddCell($row+$count+1+$i,$col+7,"$run_end_hour_r\:$min_start_end_r",$format_number);#run_stop
$worksheet->AddCell($row+$count+1+$i,$col+9,"$Dhour_r\:$Dmin_r",$format_number);#flight time
$worksheet->AddCell($row+$count+1+$i,$col+3,"$hour_start_r\:$min_start_r",$format_number);#запуск
$worksheet->AddCell($row+$count+1+$i,$col+8,"$hour_stop_r\:$min_stop_r",$format_number);#switch off
$worksheet->AddCell($row+$count+1+$i,$col+10,"$Delta_hour_r\:$Delta_min_r",$format_number);#start stop time
$worksheet->AddCell($row+$count+1+$i,$col+11,"$Fuel_cons_1_f",$format_number);#
$worksheet->AddCell($row+$count+1+$i,$col+12,"$Fuel_cons_2_f",$format_number);#
$worksheet->AddCell($row+$count+1+$i,$col+13,"$Fuel_cons_3_f",$format_number);#
$worksheet->AddCell($row+$count+1+$i,$col+14,"$Fuel_cons_4_f",$format_number);#
$worksheet->AddCell($row+$count+1+$i,$col+15,"$Fuel_cons_5_f",$format_number);#
$worksheet->AddCell($row+$count+1+$i,$col+16,"$CAS_TO_TABLE[$i]",$format_number);#
$worksheet->AddCell($row+$count+1+$i,$col+17,"$GS_TO_TABLE[$i]",$format_number);#
$worksheet->AddCell($row+$count+1+$i,$col+18,"$ATOW_TO_TABLE[$i]",$format_number);#
$worksheet->AddCell($row+$count+1+$i,$col+19,"$Route[$i]",$format_number);#
}
$template->SaveAs('START_STOP.xls');
print "--Written $i rows--"; 

my $a = <STDIN>;