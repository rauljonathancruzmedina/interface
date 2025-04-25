<?	session_start();

	include("../funciones/factorv.php");
	include("../PHPExcel/IOFactory.php");

	$op=recibe_POST('op','');
	$archivo= $_FILES['envia_file']['name'];
	$archivo_tem =$_FILES['envia_file']['tmp_name'];

	$conecta= new conector();

	$conecta->consulta("SET NAMES utf8");

	if(strlen($archivo) > 0)
	{
		if(move_uploaded_file($archivo_tem,"../interfaces/" . $archivo))
		{
			switch ($op) 
			{
				case "produccion":
					$archivo="../interfaces/" . $archivo;


					$objExcel = PHPExcel_IOFactory::load($archivo);
					$objWorksheet = $objExcel->getActiveSheet();
					$nr= $objWorksheet->getHighestRow();
					$nc = ord($objWorksheet->getHighestColumn()) - 64;

					$i=0;
					$estatus=1;
					$mensaje="";
					$hoy=date("Y-m-d");
					$capturista=$_SESSION["clave"];
					for ($row = 1; $row <= $nr; ++ $row) 
					{
						$cell = $objWorksheet->getCellByColumnAndRow(0, $row);
						$val = trim($cell->getValue());
						if($i==0)
						{
							if($val=="Clave")
							{
								$cell = $objWorksheet->getCellByColumnAndRow(1, $row);
								$val = trim($cell->getValue());
								if($val=="CURP")
								{
									$cell = $objWorksheet->getCellByColumnAndRow(3, $row);
									$val = trim($cell->getValue());
									if($val=="NSS")
									{
										for ($col = 0; $col < $nc; ++ $col)
										{
											$cell = $objWorksheet->getCellByColumnAndRow($col, $row);
											$val = trim($cell->getValue());
											switch($val)
											{
												case "Clave": $col_clave=$col; break;
												case "Folio": $col_folio=$col; break;
												case "CURP": $col_curp=$col; break;
												case "NSS": $col_nss=$col; break;
												case "Apellido Paterno": $col_apaterno=$col; break;
												case "Apellido Materno": $col_amaterno=$col; break;
												case "Nombre(s)": $col_nombre=$col; break;
												case "Fecha traspaso": $col_fec_traspaso=$col; break;
												case "Saldo": $col_saldo=$col; break;
												case "Folio AV": $col_av=$col; break;
												case "Afore Cedente": $col_afore=$col; break;
											}
										}
									}
									else
									{
										$row=$nr;
										echo "El archivo que se subió no contiene la información de la Captura de Afore";
									}
								}
								else
								{
									$row=$nr;
									echo "El archivo que se subió no contiene la información de la Captura de Afore";
								}
							}
							else
							{
								$row=$nr;
								echo "El archivo que se subió no contiene la información de la Captura de Afore";
							}
						}
						if($i==1)
						{
							$cell = $objWorksheet->getCellByColumnAndRow($col_clave, $row);
							$clave = $cell->getValue();
							if($clave != "")
							{
								$pasa=0;
								$cell = $objWorksheet->getCellByColumnAndRow($col_curp, $row);
								$curp = $cell->getValue();
								$cell = $objWorksheet->getCellByColumnAndRow($col_folio, $row);
								$folio = $cell->getValue();
								$cell = $objWorksheet->getCellByColumnAndRow($col_nss, $row);
								$nss=str_pad($cell->getValue(), 11, "0", STR_PAD_LEFT);
								$cell = $objWorksheet->getCellByColumnAndRow($col_apaterno, $row);
								$apaterno = $cell->getValue();
								$cell = $objWorksheet->getCellByColumnAndRow($col_amaterno, $row);
								$amaterno = $cell->getValue();
								$cell = $objWorksheet->getCellByColumnAndRow($col_nombre, $row);
								$nombre = $cell->getValue();

								$cell = $objWorksheet->getCellByColumnAndRow($col_fec_traspaso, $row);
								$val = $cell->getValue();
								if($val != "")
								{
									$timestamp = PHPExcel_Shared_Date::ExcelToPHP($val);
									$fecha = date("Y-m-j",$timestamp);
									$nuevafecha = strtotime ( '+1 day' , strtotime ( $fecha ) ) ;
									$fec_traspaso = date ( 'Y-m-d' , $nuevafecha );
								}
								else
									$fec_traspaso="";
								$cell = $objWorksheet->getCellByColumnAndRow($col_av, $row);
								$av = $cell->getValue();
								$cell = $objWorksheet->getCellByColumnAndRow($col_saldo, $row);
								$saldo = $cell->getValue();
								if($saldo=="")
									$saldo=0;
								$cell = $objWorksheet->getCellByColumnAndRow($col_afore, $row);
								$afore = $cell->getValue();

								$query="Select u.* from usuarios u
											inner join empleados eni on (u.clave=eni.clave) 
										where u.clave='$clave' and activo=1";
								$bd=$conecta->consulta($query,'C','c_usuario');
								if ($bd->num_rows != 0)
								{
									if($curp != "")
									{
										if($fec_traspaso != "")
										{
											if($nombre != "")
											{
												$bd=$conecta->consulta("Select id_afore from afores where afore='$afore'");
												if ($bd->num_rows != 0)
												{
													$tb=mysqli_fetch_array($bd);
													$afore=$tb["id_afore"];
													if($saldo !=0)
														$pasa=1;
													else
													{
														if($mensaje=="")
															$mensaje="- El producto con $curp no tiene saldo\n";
														else
															$mensaje.="- El producto con $curp no tiene saldo\n";
													}
												}
												else
												{
													if($mensaje=="")
														$mensaje="- El producto con $curp no tiene afore cedente correcta\n";
													else
														$mensaje.="- El producto con $curp no tiene afore cedente correcta\n";
												}
											}
											else
											{
												if($mensaje=="")
													$mensaje="- El producto con $curp no tiene nombre del cliente\n";
												else
													$mensaje.="- El producto con $curp no tiene nombre del cliente\n";
											}
										}
										else
										{
											if($mensaje=="")
												$mensaje="- El producto con $curp no tiene fecha de traspaso\n";
											else
												$mensaje.="- El producto con $curp no tiene fecha de traspaso\n";
										}
									}
									else
									{
										if($mensaje=="")
											$mensaje="- Existe un registro sin $curp\n";
										else
											$mensaje.="- Existe un registro sin $curp\n";
									}
								}
								else
								{
									if($mensaje=="")
										$mensaje="- La clave $clave del producto con CURP $curp\nno existe o no está activo\n";
									else
										$mensaje.="- La clave $clave del producto con CURP $curp\nno existe o no está activo\n";
								}
								if($pasa==1)
								{
									$bd=$conecta->consulta("Select * from produccion where curp='$curp'");
									if ($bd->num_rows == 0)
									{
										$query="Insert into produccion values('$curp','$folio','$nss','$hoy','$fec_traspaso','$nombre','$apaterno','$amaterno','$capturista','$clave',$afore,'$av',$saldo,1,1,'',0)";
										$conecta->consulta($query,'I','i_prod');
									}
									else
									{
										if($mensaje=="")
											$mensaje="- Ya existe un registro con el CURP $curp.\nIngrésalo directamente en la página de captura";
										else
											$mensaje.="- Ya existe un registro con el CURP $curp.\nIngrésalo directamente en la página de captura";
									}
								}
							}
						}
						$i=1;
					}
					if($mensaje != "")
						echo $mensaje . "\n\nESTOS PRODUCCTOS NO SUBIERON AL SISTEMA";
					else
						echo "Todos los registros subieron correctamente";
				break;		
			}
		}
		else
			echo "No pudo subir el archivo al servidor. Inténtalo nuevamente...";
	}
	else
		echo "No se ha cargado el archivo... Intenta de nuevo";
?>