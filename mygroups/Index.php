<?php
session_start();
error_reporting(E_ALL|E_STRICT);
ini_set("display_errors", 1);

require_once("Settings.php");
require_once("AuthHelper.php");
require_once("Token.php");

//check for token in session first time in
if (!isset($_SESSION[Settings::$tokenCache])) {
  //set the isadd flag if necessary
  if (isset($_GET["addin"]))
    $_SESSION[Settings::$isAddin] = True;
  else
    $_SESSION[Settings::$isAddin] = False;

  //redirect to login page
  header("Location:Login.php");
}
else {
  //get addin value
  $isaddin = $_SESSION[Settings::$isAddin];

  //use the refresh token to get a new access token
  $token = AuthHelper::getAccessTokenFromRefreshToken($_SESSION[Settings::$tokenCache]);

  //perform a REST query for the users modern groups
  $request = curl_init(Settings::$unifiedAPIEndpoint . "me/joinedgroups");
  curl_setopt($request, CURLOPT_HTTPHEADER, array(
    "Authorization: Bearer " . $token->accessToken,
    "Accept: application/json"));
  curl_setopt($request, CURLOPT_RETURNTRANSFER, true);
  $response = curl_exec($request);

  //parse the json into oci_fetch_object
  $groups = json_decode($response, true);
  $apiRoot = $groups["@odata.context"];
  $apiRoot = substr($apiRoot, 0, strrpos($apiRoot, "/"));
  $_SESSION[Settings::$apiRoot] = $apiRoot;
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

    });
  }
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
        <table class="table table-striped table-hover">
          <thead>
            <tr>
              <th>Group Name</th>
              <th>Email</th>
            </tr>
          </thead>
          <tbody>
            <?php foreach ($groups["value"] as $group) { ?>
              <tr>
                <td><a href="Detail.php?id=<?php echo $group["objectId"] ?>"><?php echo $group["displayName"] ?></a></td>
                <td><a href="mailto:<?php echo $group["EmailAddress"] ?>"><?php echo $group["EmailAddress"] ?></a></td>
              </tr>
            <?php } ?>
          </tbody>
        </table>
      </div>
    </div>
  </div>
</body>
</html>
