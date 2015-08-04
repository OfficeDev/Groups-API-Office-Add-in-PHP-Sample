<?php
session_start();
error_reporting(E_ALL|E_STRICT);
ini_set("display_errors", 1);

require_once("Settings.php");
require_once("AuthHelper.php");
require_once("Token.php");

//get addin value
$isaddin = $_SESSION[Settings::$isAddin];

//check for authorization code in url parameter
if (isset($_GET["code"])) {
  //use the authorization code to get access token for the unified API
  $token = AuthHelper::getAccessTokenFromCode($_GET["code"]);
  if (isset($token->refreshToken)) {
    $_SESSION[Settings::$tokenCache] = $token->refreshToken;
    header("Location:Index.php");
  }
}
 ?>

<html>
<head>
  <title>Login</title>
  <link rel="stylesheet" href="css/bootstrap.min.css">
  <script type="text/javascript" src="scripts/jquery-1.10.2.min.js"></script>
  <script type="text/javascript" src="scripts/bootstrap.min.js"></script>
  <script type="text/javascript">
    function login() {
      <?php
      echo "window.location = '" . AuthHelper::getAuthorizationUrl() . "';";
       ?>
    }
  </script>
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
      <h1>My Modern Groups</h1>
    </div>
    <div class="row">
      <div class="col-sm-12">
        <div class="panel panel-default">
          <div class="panel-body">
            <h3>Login Required</h3>
            <p>My Modern Groups is a web application that queries Office 365 to display the modern groups you are a member of. To query Office 365, you must first login with Office 365 credientals and then grant this application access to query group data.</p>
            <button type="button" name="button" onclick="login()" class="btn btn-primary btn-block">Login with Office 365</button>
          </div>
        </div>
      </div>
    </div>
  </div>
</body>
</html>
