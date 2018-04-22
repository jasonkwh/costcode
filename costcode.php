<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="styles/bootstrap.min.css">
    <link rel="stylesheet" href="styles/notify.css">
    <link rel="stylesheet" href="styles/prettify.css">
    <link rel="stylesheet" href="styles/fontawesome-all.min.css">
    <script src="scripts/jquery.min.js"></script>
    <script src="scripts/popper.min.js"></script>
    <script src="scripts/bootstrap.min.js"></script>
    <script src="scripts/notify.js"></script>
    <script src="scripts/prettify.js"></script>
    <title>COST CODE</title>
    <script>
        function ajaxfileupload() {
            var file_data = $('#fileToUpload').prop('files')[0];
            var form_data = new FormData();
            form_data.append('fileToUpload', file_data);
            $.ajax({
                url: 'costcodebackend.php', // point to server-side PHP script
                dataType: 'text',  // what to expect back from the PHP script, if anything
                cache: false,
                contentType: false,
                processData: false,
                data: form_data,
                type: 'post',
                success: function(data){
                    if(data=="error") {
                        $.notify("<span class='fas fa-exclamation-circle'></span> ERROR: Please contact <a href=\"https://github.com/jasonkwh\">Jason Huang</a>.", {type:"danger",close:true,delay:3000});
                    } else {
                        location.href = data;
                    }
                }
            });
        }
    </script>
</head>
<body>
<div class="container">
    <div class="row justify-content-center" style="margin:10px"><h4>COST CODE v1.0</h4></div>
    <div class="row justify-content-center">
        <div class="card" style="width: 30rem;">
            <div class="card-body">
                <h5 class="card-title">Open an Excel file</h5>
                <div class="row">
                <div class="col-8">
                <p class="card-text"><input id="fileToUpload" name="fileToUpload" type="file"></p>
                </div>
                <div class="col-4">
                <button type="button" class="btn btn-dark" style="margin-top:-10px" onclick="ajaxfileupload()"><i class="far fa-file-excel"></i>&nbsp;&nbsp;Upload</button>
                </div>
                </div>
            </div>
        </div>
    </div>
    <div class="row justify-content-center" style="margin:10px"><p>made with <i class="fab fa-gratipay" style="color:red"></i> by <a href="https://github.com/jasonkwh">Jason Huang</a> <?php echo date("Y"); ?>. Love BB forever!</p></div>
</div>
</body>
</html>