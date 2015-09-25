<?php
date_default_timezone_set('America/Montevideo');
require_once 'PHPExcel.php';

$filename = uniqid(rand(), true) . '.xls';

$possible_sent = true;

// Check if all args are here.
if (($_POST["schoolname"] == null) ||
    ($_POST["email"] == null) ||
    ($_POST["moderatorname"] == null) ||
    ($_POST["number"] == null)) {
  $possible_sent = false;
};

for ($i=1; $i<=34; $i++) {
  if ($_POST['participant'.$i] == null) {
    $possible_sent = false;
    break;
  }
}

if ($possible_sent == false) {
  header("Location: /?register=false#register");
}
else {
  $exporter = new PHPExcel();
  $exporter->getProperties()
              ->setCreator("symbiosis15")
              ->setLastModifiedBy("symbiosis15 web")
              ->setTitle("Symbiosis15 registration form")
              ->setSubject("Symbiosis15 registration form")
              ->setDescription("Symbiosis15 registration form");

  $exporter->setActiveSheetIndex(0)
              ->setCellValue('A1', 'Input')
              ->setCellValue('B1', 'Value')
              ->setCellValue('A2', 'School name')
              ->setCellValue('B2', $_POST['schoolname'])
              ->setCellValue('A3', 'Name of teacher')
              ->setCellValue('B3', $_POST['moderatorname'])
              ->setCellValue('A4', 'Number')
              ->setCellValue('B4', $_POST['number'])
              ->setCellValue('A5', 'Email')
              ->setCellValue('B5', $_POST['email']);

  for ($i=1; $i<=34; $i++) {
    $exporter->setActiveSheetIndex(0)
             ->setCellValue('A'.($i+5), 'Participant '.$i)
             ->setCellValue('B'.($i+5), $_POST['participant'.$i]);
  }

  foreach(range('A','B') as $columnID) {
      $exporter->getActiveSheet()->getColumnDimension($columnID)
          ->setAutoSize(true);
  }

  $exporter->getActiveSheet()->setTitle('Data');
  $exporter->setActiveSheetIndex(0);
  $objWriter = PHPExcel_IOFactory::createWriter($exporter, 'Excel2007');
  $objWriter->save("register_xls/".$filename);

  $url = 'https://api.sendgrid.com/';
  $user = 'ignauy';
  $pass = 'symbiosis15';
  $filePath = 'register_xls';

  $params = array(
      'api_user'  => $user,
      'api_key'   => $pass,
      'to'        => 'vipulsharma936@gmail.com',
      'cc'        => 'nachoel01@gmail.com',
      'subject'   => 'New register from page',
      'html'      => 'XLS file attached.',
      'text'      => 'XLS file attached.',
      'from'      => 'contact@symbiosis15.net',
      'files[register.xls]' => '@'.$filePath.'/'.$filename
    );

  $request =  $url.'api/mail.send.json';

  $session = curl_init($request);

  curl_setopt ($session, CURLOPT_POST, true);
  curl_setopt ($session, CURLOPT_POSTFIELDS, $params);
  curl_setopt($session, CURLOPT_HEADER, false);
  curl_setopt($session, CURLOPT_SSLVERSION, CURL_SSLVERSION_TLSv1_2);
  curl_setopt($session, CURLOPT_RETURNTRANSFER, true);
  $response = curl_exec($session);
  curl_close($session);

  header("Location: /?register=true#postregister");
};
?>
