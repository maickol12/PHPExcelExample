<?php


  include_once("PHPExcel.php");
  $excel = new PHPExcel();
  //Usamos el worsheet por defecto
  $sheet = $excel->getActiveSheet();
  //Agregamos un texto a la celda A1
  $sheet->setCellValue('A1', 'Prueba');
  //Damos formato o estilo a nuestra celda
  $sheet->getStyle('A1')->getFont()->setName('Tahoma')->setBold(true)->setSize(8);
  $sheet->getStyle('A1')->getBorders()->applyFromArray(array('allBorders' => 'thin'));
  $sheet->getStyle('A1')->getAlignment()->setVertical('center')->setHorizontal('center');
  $sheet->getStyle('A1')->applyFromArray(
    array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => '202124')
        )
    )
  );
  $styleArray = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 10,
        'name'  => 'Verdana'
    ));
  $sheet->getStyle('A1')->applyFromArray($styleArray);

  $sheet->setCellValue('B1', 'PHPExcel');

  //usamos los mismos estilos de A1
  $sheet->getStyle('B1')->getFont()->setName('Tahoma')->setBold(true)->setSize(8);
  $sheet->getStyle('B1')->getBorders()->applyFromArray(array('allBorders' => 'thin'));
  $sheet->getStyle('B1')->getAlignment()->setVertical('center')->setHorizontal('center');
  //exportamos nuestro documento
  $writer = new PHPExcel_Writer_Excel5($excel);
  $writer->save('prueba.xls');

  header('Content-Type: application/csv');
  header('Content-Disposition: attachment; filename=prueba.xls');
  header('Pragma: no-cache');
  readfile("prueba.xls");
?>
