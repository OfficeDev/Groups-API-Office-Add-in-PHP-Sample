<?php
session_start();
error_reporting(E_ALL|E_STRICT);
ini_set("display_errors", 1);

require_once("Settings.php");
require_once("AuthHelper.php");
require_once("Token.php");

//check for token in session first time in
if (!isset($_SESSION[Settings::$tokenCache])) {
  //redirect to login page
  header("Location:Login.php");
}
else {
  //check for id and apiRoot
  if (!isset($_SESSION[Settings::$apiRoot]) || !isset($_GET["id"])) {
    //redirect back to Index.php
    header("Location:Index.php");
  }
  else {
    //get addin value
    $isaddin = $_SESSION[Settings::$isAddin];

    //get the apiRoot from session
    $apiRoot = $_SESSION[Settings::$apiRoot];

    //get the id from url parameter
    $id = $_GET["id"];

    //use the refresh token to get a new access token
    $token = AuthHelper::getAccessTokenFromRefreshToken($_SESSION[Settings::$tokenCache]);

    //perform a REST query for the users modern groups
    $request = curl_init($apiRoot . "/groups/" . $id . "/members");
    curl_setopt($request, CURLOPT_HTTPHEADER, array(
      "Authorization: Bearer " . $token->accessToken,
      "Accept: application/json"));
    curl_setopt($request, CURLOPT_RETURNTRANSFER, true);
    $response = curl_exec($request);

    //parse the json into oci_fetch_object
    $members = json_decode($response, true);
  }
}
?>
<html>
<head>
  <title>My Groups</title>
  <link rel="stylesheet" href="css/bootstrap.min.css">
  <script type="text/javascript" src="scripts/jquery-1.10.2.min.js"></script>
  <script type="text/javascript" src="scripts/bootstrap.min.js"></script>
  <?php if ($isaddin) { ?>
  <script type="text/javascript" src="//appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <script type="text/javascript">
  //initialize Office on each page if add-in
  Office.initialize = function(reason) {
    $(document).ready(function() {
      var data = JSON.parse(excelData);
      var officeTable = new Office.TableData();

      //build headers
      var headers = new Array("Name", "Email", "Job Title", "Department");
      officeTable.headers = headers;

      //add data
      for (var i = 0; i < data.value.length; i++) {
        officeTable.rows.push([data.value[i].displayName,
          data.value[i].mail,
          data.value[i].jobTitle,
          data.value[i].department]);
      }

      //add the table to Excel
      Office.context.document.setSelectedDataAsync(officeTable, { coercionType: Office.CoercionType.Table }, function (asyncResult) {
        //check for error
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          $("#error").show();
        }
        else {
          $("#success").show();
        }
      });
    });
  }

  var excelData = '<?php echo $response ?>';
  </script>
  <?php } ?>
</head>
<body>
  <div class="container">
    <div class="page-header page-header-inverted">
      <h1><a href="index.php">My Modern Groups</a></h1>
    </div>
    <div class="row">
      <div class="col-sm-12">
        <div class="alert alert-success" role="alert" style="display: none;" id="success">
          SUCCESS: Update to Excel succeeded!
        </div>
        <div class="alert alert-danger" role="alert" style="display: none;" id="error">
          ERROR: Update to Excel failed!
        </div>
        <table class="table table-striped table-hover">
          <thead>
            <tr>
              <th>Name</th>
              <th>Email</th>
              <th>Title</th>
              <th>Department</th>
            </tr>
          </thead>
          <tbody>
            <?php foreach ($members["value"] as $member) { ?>
              <tr>
                <td><?php echo $member["displayName"] ?></td>
                <td><a href="mailto:<?php echo $member["mail"] ?>"><?php echo $member["mail"] ?></a></td>
                <td><?php echo $member["jobTitle"] ?></td>
                <td><?php echo $member["department"] ?></td>
              </tr>
            <?php } ?>
          </tbody>
        </table>
      </div>
    </div>
  </div>
</body>
</html>
