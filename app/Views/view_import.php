<!doctype html>
<html lang="en">

<head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">

    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.3/css/jquery.dataTables.css">
    <style>
        .tabledata{
            overflow-x: scroll;
        }
    </style>

    <title>Hello, world!</title>
</head>

<body>
    <nav class="navbar navbar-expand-lg navbar-dark bg-success fixed-top" style="box-shadow: 0 8px 10px -5px rgb(0 0 0 / 16%), 0 16px 24px 2px rgb(0 0 0 / 11%), 0 6px 30px 5px rgb(0 0 0 / 10%);">
        <div class="container">
            <a class="navbar-brand" href="#">Daftar Nama Peserta</a>
            <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav" aria-controls="navbarNav" aria-expanded="false" aria-label="Toggle navigation">
                <span class="navbar-toggler-icon"></span>
            </button>
            <div class="collapse navbar-collapse" id="navbarNav">
                <ul class="navbar-nav">
                    <li class="nav-item">
                        <a class="nav-link active" aria-current="page" href="<?= site_url('Import/index') ?>">Home</a>
                    </li>
                </ul>
            </div>
        </div>
    </nav>

    <div class="container mt-5">
        <div class="row justify-content-between">
            <div class="col-md-4 col-12">
                <div class="card mt-4" style="box-shadow: 0 8px 10px -5px rgb(0 0 0 / 16%), 0 16px 24px 2px rgb(0 0 0 / 11%), 0 6px 30px 5px rgb(0 0 0 / 10%);">
                    <div class="card-header bg-success">
                        <h4 class="text-center text-white">Import Data Excel</h4>
                    </div>
                    <div class="card-body p-4 border border-success">
                        <?= form_open_multipart("import/upload") ?>
                        <div class="form_group row">
                            <?php if (!empty(session()->getFlashdata('error'))) { ?>
                                <div class="alert alert-danger alert-dismissible fade show mt-1 px=0" role="alert">
                                    <strong>Maaf!</strong> <?php echo session()->getFlashdata('error'); ?>
                                    <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                                </div>
                            <?php } ?>
                            <div class="col-md-12 col-12">
                                <label class="col-form-label">Import Data</label>
                                <input type="file" name="excel" class="form-control">
                                <button type="submit" class="btn btn-success mt-2 w-100">Upload</button>
                            </div>
                            <div class="col-md-12 col-12">
                                <a href="<?= base_url("import/download") ?>" class="btn btn-outline-success mt-2 w-100">Dowload Format</a>
                            </div>
                        </div>
                        <?= form_close() ?>
                    </div>
                </div>
            </div>

            <div class="col-md-8 col-12">
                <div class="card mt-4" style="box-shadow: 0 8px 10px -5px rgb(0 0 0 / 16%), 0 16px 24px 2px rgb(0 0 0 / 11%), 0 6px 30px 5px rgb(0 0 0 / 10%);">
                    <div class="card-header bg-success">
                        <h4 class="text-center text-white mb-3">Daftar Nama Peserta Bimbingan Teknis dan Pendampingan   </h4>
                    </div>

                    <div class="card-body border border-success">
                        <?php if (!empty(session()->getFlashdata('sukses'))) {  ?>
                            <div class="alert alert-success alert-dismissible fade show mt-1 px=0" role="alert">
                                <strong>PESAN</strong><br> <?php echo session()->getFlashdata('sukses'); ?>
                                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                            </div>
                        <?php } ?>

                        <a href="<?= base_url("Import/export") ?>" class="btn btn-success">Export Excel</a>

                        <div class="table-responsive mt-3 tabledata" >
                            <table class="table table-striped table-bordered" id="myTable">
                                <thead class="table-light">
                                    <tr>
                                        <th>No</th>
                                        <th >Nama</th>
                                        <th >NIP</th>
                                        <th> Unit kerja</th>
                                        <th> Kab/Kota</th>
                                        <th>Email</th>
                                        <th>No_HP</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <?php foreach ($users as $key => $user) : ?>
                                        <tr>
                                            <td><?= $key+1 ?></td>
                                            <td><?= $user["nama"] ?></td>
                                            <td><?= $user["nip"] ?></td>
                                            <td><?= $user["unit_kerja"] ?></td>
                                            <td><?= $user["kabkota"] ?></td>
                                            <td><?= $user["email"] ?></td>
                                            <td><?= $user["no_hp"]?></td>
                                            <td>
                                                <a href="<?= base_url("Import/delete/" . $user["id"]) ?>" class="btn btn-danger btn-sm">Hapus</a>
                                            </td>
                                        </tr>
                                    <?php endforeach; ?>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js" integrity="sha256-/xUj+3OJU5yExlq6GSYGSHk7tPXikynS7ogEvDej/m4=" crossorigin="anonymous"></script>

    <script type="text/javascript" charset="utf8" src="https://cdn.datatables.net/1.11.3/js/jquery.dataTables.js"></script>


    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>

    <script>
        $(document).ready(function() {
            $('#myTable').DataTable();
        });
    </script>

    <script>
        window.setTimeout(function() {
            $(".alert").fadeTo(500, 0).slideUp(500, function() {
                $(this).remove();
            });
        }, 5000);
    </script>
</body>

</html>