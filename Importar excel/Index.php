<?php

$idPaisFiltrado = $_GET['p'];

if (isset($_GET['r'])) {
  $param = explode('-', $_GET['r']);
  $id_marca = base64_decode($param[1]);
}

$marca = new MarcasController();
$marca->update_visitas($id_marca);
$marca = $marca->read($id_marca, false);

$años_experiencia = date('Y') - $marca['year_fundacion'];

$franquicias = new FranquiciasController();
$franquicias = $franquicias->read($id_marca, $idPaisFiltrado);

$fondos_comercio = new FondosComercioController();
$fondos_comercio = $fondos_comercio->read($id_marca, $idPaisFiltrado);

$paises = new PaisesController();
$paises = $paises->read_contacto();

$negocios = new MarcasController();
$negocios = $negocios->get_negocios($id_marca, $idPaisFiltrado);

$publicidad = new PublicidadesController();
$publicidad = $publicidad->read(4);

if (count($franquicias) > 0) {
  $contacto = $franquicias[0];
} elseif (count($fondos_comercio) > 0) {
  $contacto = $fondos_comercio[0];
}

?>

<section>
    <div class="row valign-wrapper">
        <div class="col s12 m6 l3 xl3 valign-wrapper">
            <img src="<?= SERVER_URL . 'back/' . $marca['path_logo'] ?>" class="img-responsive"
                style="margin: auto; height: 100px">
        </div>
        <div class="col s12 m6 l3 xl3 center" style="background-color: #75A8B4; color: white;">
            <h1><?= $años_experiencia ?></h1>
            <h5>años de experiencia</h5>
        </div>
        <div class="col s12 m6 l3 xl3 center" style="background-color: #95BDC6;">
            <h1><?= $marca['locales_propios'] ?></h1>
            <h5>locales propios</h5>
        </div>
        <div class="col s12 m6 l3 xl3 center" style="background-color: #47727C; color: white;">
            <h1><?= $marca['franquicias_operativas'] ?></h1>
            <h5>franquicias operativas</h5>
        </div>
    </div>
    <div class="row encabezado"
        style="background-image: url(<?= SERVER_URL . 'back/' . $marca['path_img_cabecera'] ?>)">
        <div class="col s12 m10 l5 xl5 right">
            <h5 class="inversion center"><?= $marca['nombre_marca'] ?></h5>
            <p class="descripcion"><?= $marca['descripcion'] ?></p>
        </div>
    </div>
</section>

<section>
    <div class="row" style="margin-top: 20px">
        <div class="col s12 m12 l8 xl8">

            <?php
      $cantidad_tabs = count($franquicias) + count($fondos_comercio);
      if ($cantidad_tabs > 0) {
        echo '
          <ul class="tabs">
          ';
        $col = 12 / $cantidad_tabs;
        foreach ($franquicias as $key => $franquicia) {
          echo '
            <li class="tab col s' . $col . '"><a class="active" href="#franquicia' . $franquicia['id_franquicia'] . '" style="color: white;">' . $franquicia['det_tipo_negocio'] . '</a></li>
            ';
        }
        foreach ($fondos_comercio as $key => $fondo) {
          echo '
            <li class="tab col s' . $col . '"><a class="active" href="#fondo_comercio' . $fondo['id_fondo_comercio'] . '" style="color: white;">' . $fondo['det_tipo_negocio'] . '</a></li>
            ';
        }
        echo '
          </ul>
          ';

        foreach ($franquicias as $key => $franquicia) {
          echo '
            <div id="franquicia' . $franquicia['id_franquicia'] . '" class="col s12">

              <div class="card">
                <div class="card-content datos">

                  <div class="row">

                    <h5 class="titulo-principal">Franquicia: ' . $franquicia['det_tipo_negocio'] . '</h5>
                    <h6 class="subtitulo-principal">Costo de inversión: ' . $franquicia['pre_inversion_total'] . ' ' . $franquicia['moneda'] . $franquicia['inversion_total'] . '</h6>

                    <div class="col s12 m12 l6 xl6">

                      <h5 class="titulo-tabla">Datos de la empresa</h5>

                      <table class="highlight tabla-datos">
                        <tbody>';

          if (!empty($marca['razon_social'])) {
            echo '
                          <tr>
                            <td><span>Razón social: </span>' . $marca['razon_social'] . '</td>
                          </tr>';
          }
          echo '
                          <tr>
                            <td><span>Rubro/s: </span>' . $marca['rubros'] . '</td>
                          </tr>
                          <tr>
                            <td><span>País de origen: </span>' . $marca['det_pais'] . '</td>
                          </tr>
                          <tr>
                            <td><span>Paises en los que actúa: </span>' . $marca['paises_actua'] . '</td>
                          </tr>
                          <tr>
                            <td><span>Año de fundación: </span>' . $marca['year_fundacion'] . '</td>
                          </tr>';

          if ($marca['primera_franquicia'] != '0000-00-00') {
            echo '
                            <tr>
                              <td><span>Constitución primera franquicia: </span>' . $marca['primera_franquicia'] . '</td>
                            </tr>
                            ';
          }

          if ($marca['franquicias_last_year'] != '0') {
            echo '
                            <tr>
                              <td><span>Franquicias abiertas en el último año: </span>' . $marca['franquicias_last_year'] . '</td>
                            </tr>
                            ';
          }

          echo
          '</tbody>
                      </table>

                      <h5 class="titulo-tabla">Datos del local</h5>

                      <table class="highlight tabla-datos">
                        <tbody>
                          <tr>
                            <td><span>Dimensiones mínimas: </span>' . $franquicia['dimensiones'] . '</td>
                          </tr>
                          <tr>
                            <td><span>Población mínima: </span>' . $franquicia['poblacion_minima'] . '</td>
                          </tr>
                          <tr>
                            <td><span>Ubicación preferible: </span>' . $franquicia['ubicacion_preferible'] . '</td>
                          </tr>
                          <tr>
                            <td><span>Local llave en mano: </span>' . $franquicia['llave_mano'] . '</td>
                          </tr>
                        </tbody>
                      </table>

                      <h5 class="titulo-tabla">Datos económicos</h5>

                      <table class="highlight tabla-datos">
                        <tbody>
                          <tr>
                            <td><span>Canon de ingreso: </span>' . $franquicia['pre_canon_ingreso'] . ' ' . $franquicia['moneda'] . ' ' . number_format($franquicia['canon_ingreso'], 0, ',', '.') . '</td>
                          </tr>
                          <tr>
                            <td><span>Costo de obra - variable: </span>' . $franquicia['pre_costo_obra'] . ' ' . $franquicia['moneda'] . ' ' . number_format($franquicia['costo_obra'], 0, ',', '.') . '</td>
                          </tr>
                          <tr>
                            <td><span>Inversión total aproximada: </span>' . $franquicia['pre_inversion_total'] . ' ' . $franquicia['moneda'] . ' ' . number_format($franquicia['inversion_total'], 0, ',', '.') . '</td>
                          </tr>
                        </tbody>
                      </table>';

          echo '
                    </div>

                    <div class="col s12 m12 l6 xl6">

                      <h5 class="titulo-tabla">Datos económicos</h5>

                      <table class="highlight tabla-datos">
                        <tbody>';

          if (!empty($franquicia['lanzamiento_publicidad']) && $franquicia['lanzamiento_publicidad'] != '0') {
            echo '<tr>
                          <td><span>Lanzamiento de Publicidad en apertura: </span>' . $franquicia['pre_lanzam_publicidad'] . ' ' . $franquicia['moneda'] . ' ' . $franquicia['lanzamiento_publicidad'] . '</td>
                        </tr>';
          }

          if (!empty($franquicia['stock_inicial']) && $franquicia['stock_inicial'] != '0') {
            echo '<tr>
                    <td><span>Stock Inicial: </span>' . $franquicia['pre_stock_inicial'] . ' ' . $franquicia['moneda'] . ' ' . $franquicia['stock_inicial'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['regalias']) && $franquicia['regalias'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Regalías: </span>' . $franquicia['regalias'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['canon_publicidad']) && $franquicia['canon_publicidad'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Canon de publicidad: </span>' . $franquicia['canon_publicidad'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['facturacion_promedio']) && $franquicia['facturacion_promedio'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Facturación promedio mensual: </span>' . $franquicia['facturacion_promedio'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['promedio_empleados']) && $franquicia['promedio_empleados'] != 0) {
            echo '<tr>
                    <td><span>Promedio de empleados por local: </span>' . $franquicia['promedio_empleados'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['financiacion_disponible']) && $franquicia['financiacion_disponible'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Financiación disponible: </span>' . $franquicia['financiacion_disponible'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['duracion_contrato']) && $franquicia['duracion_contrato'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Duración del contrato: </span>' . $franquicia['duracion_contrato'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['recupero_inversion']) && $franquicia['recupero_inversion'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Recupero de la inversión: </span>' . $franquicia['recupero_inversion'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['publicidad_corporativa']) && $franquicia['publicidad_corporativa'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Publicidad corporativa: </span>' . $franquicia['publicidad_corporativa'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['tipo_franquiciado']) && $franquicia['tipo_franquiciado'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Tipo de franquiciado: </span>' . $franquicia['tipo_franquiciado'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['proceso_seleccion']) && $franquicia['proceso_seleccion'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Proceso de selección del franquiciado: </span>' . $franquicia['proceso_seleccion'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['lugar_y_duracion']) && $franquicia['lugar_y_duracion'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Duración del entrenamiento: </span>' . $franquicia['lugar_y_duracion'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['manual']) && $franquicia['manual'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Manual de operaciones: </span>' . $franquicia['manual'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['rentabilidad']) && $franquicia['rentabilidad'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Rentabilidad: </span>' . $franquicia['rentabilidad'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['socio_camara']) && $franquicia['socio_camara'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Socio de cámara de franquicias: </span>' . $franquicia['socio_camara'] . '</td>
                  </tr>';
          }

          if (!empty($franquicia['datos_adicionales'])) {
            echo '<tr>
                    <td><span>Datos adicionales: </span>' . $franquicia['datos_adicionales'] . '</td>
                  </tr>';
          }

          echo '
                </tbody>
              </table>
            </div>

          </div>';

          if (!empty($franquicia['fotos'])) {

            $fotos = explode(',', $franquicia['fotos']);

            echo '
                    <div class="row" style="width: 100%; margin: 10px auto">
                      <div class="mygallery" id="mygallery">';

            foreach ($fotos as $key => $foto) {

              $data_foto = explode('|', $foto);

              echo '
                        <a>
                          <img class="img-franquicias" src="' . SERVER_URL . 'back/' . $data_foto[1] . '" data-id-foto="' . $data_foto[0] . '" />
                        </a>
                      ';
            }

            echo '
                      </div>
                    </div>';
          }

          echo '
                </div>
              </div>

            </div>
            ';
        }

        foreach ($fondos_comercio as $key => $fondo) {
          echo '
            <div id="fondo_comercio' . $fondo['id_fondo_comercio'] . '" class="col s12">

              <div class="card">
                <div class="card-content datos">

                  <div class="row">

                    <h5 class="titulo-principal">Fondo de comercio: ' . $fondo['det_tipo_negocio'] . '</h5>
                    <h6 class="subtitulo-principal">Precio de venta: ' . $fondo['moneda'] . $fondo['precio_venta'] . '</h6>

                    <div class="col s12 m12 l6 xl6">

                      <h5 class="titulo-tabla">Datos de la empresa</h5>

                      <table class="highlight tabla-datos">
                        <tbody>';

          if (!empty($marca['razon_social'])) {
            echo '
                          <tr>
                            <td><span>Razón social: </span>' . $marca['razon_social'] . '</td>
                          </tr>';
          }

          echo '
                          <tr>
                            <td><span>Rubro/s: </span>' . $marca['rubros'] . '</td>
                          </tr>
                          <tr>
                            <td><span>País de origen: </span>' . $marca['det_pais'] . '</td>
                          </tr>
                          <tr>
                            <td><span>Paises en los que actúa: </span>' . $marca['paises_actua'] . '</td>
                          </tr>
                          <tr>
                            <td><span>Año de fundación: </span>' . $marca['year_fundacion'] . '</td>
                          </tr>';

          if ($marca['primera_franquicia'] != '0000-00-00') {
            echo '
                          <tr>
                            <td><span>Constitución primera franquicia: </span>' . $marca['primera_franquicia'] . '</td>
                          </tr>';
          }

          if ($marca['franquicias_last_year'] != '0') {
            echo '
                          <tr>
                            <td><span>Franquicias abiertas en el último año: </span>' . $marca['franquicias_last_year'] . '</td>
                          </tr>';
          }

          echo '
                        </tbody>
                      </table>

                      <h5 class="titulo-tabla">Datos del local</h5>

                      <table class="highlight tabla-datos">
                        <tbody>
                          <tr>
                            <td><span>Dimensiones: </span>' . $fondo['dimensiones'] . '</td>
                          </tr>
                          <tr>
                            <td><span>Local: </span>' . $fondo['alquilado_propio'] . '</td>
                          </tr>';

          if ($fondo['alquilado_propio'] == 'Alquilado') {
            echo '
                            <tr>
                              <td><span>Monto del alquiler: </span>' . $fondo['monto_alquiler'] . '</td>
                            </tr>
                            ';
          }

          echo '
                        </tbody>
                      </table>';

          echo '
                    </div>

                    <div class="col s12 m12 l6 xl6">

                      <h5 class="titulo-tabla">Datos económicos</h5>

                      <table class="highlight tabla-datos">
                        <tbody>';
          if (!empty($fondo['ventas_last_year']) && $fondo['ventas_last_year'] != 0) {
            echo '<tr>
                    <td><span>Ventas en el último año: </span>' . $fondo['moneda'] . ' ' . number_format($fondo['ventas_last_year'], 0, ',', '.') . '</td>
                  </tr>';
          }

          if (!empty($fondo['costo_venta']) && $fondo['costo_venta'] != 0) {
            echo '<tr>
                    <td><span>Costo de venta: </span>' . $fondo['moneda'] . ' ' . number_format($fondo['costo_venta'], 0, ',', '.') . '</td>
                  </tr>';
          }

          if (!empty($fondo['mano_obra']) && $fondo['mano_obra'] != 0) {
            echo '<tr>
                    <td><span>Costo de mano de obra: </span>' . $fondo['moneda'] . ' ' . number_format($fondo['mano_obra'], 0, ',', '.') . '</td>
                  </tr>';
          }

          if (!empty($fondo['costo_operativo']) && $fondo['costo_operativo'] != 0) {
            echo '<tr>
                    <td><span>Costo operativo: </span>' . $fondo['moneda'] . ' ' . number_format($fondo['costo_operativo'], 0, ',', '.') . '</td>
                  </tr>';
          }

          if (!empty($fondo['rentabilidad_promedio']) && $fondo['rentabilidad_promedio'] != 0) {
            echo '<tr>
                    <td><span>Rentabilidad promedio: </span>' . $fondo['moneda'] . ' ' . number_format($fondo['rentabilidad_promedio'], 0, ',', '.') . '</td>
                  </tr>';
          }

          if (!empty($fondo['inventario_bienes']) && $fondo['inventario_bienes'] != 0) {
            echo '<tr>
                    <td><span>Inventario de bienes: </span>' . $fondo['moneda'] . ' ' . number_format($fondo['inventario_bienes'], 0, ',', '.') . '</td>
                  </tr>';
          }

          if (!empty($fondo['financiacion_disponible']) && $fondo['financiacion_disponible'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Financiación disponible: </span>' . $fondo['financiacion_disponible'] . '</td>
                  </tr>';
          }

          if (!empty($fondo['otros_datos']) && $fondo['otros_datos'] != 'Sin datos') {
            echo '<tr>
                    <td><span>Otros datos de interés: </span>' . $fondo['otros_datos'] . '</td>
                  </tr>';
          }

          echo '
                </tbody>
              </table>
            </div>

          </div>';

          if (!empty($fondo['fotos'])) {

            $fotos = explode(',', $fondo['fotos']);

            echo '
                    <div class="row" style="width: 100%; margin: 10px auto">
                      <div class="mygallery" id="mygallery">';

            foreach ($fotos as $key => $foto) {

              $data_foto = explode('|', $foto);

              echo '
                        <a>
                          <img class="img-franquicias" src="' . SERVER_URL . 'back/' . $data_foto[1] . '" data-id-foto="' . $data_foto[0] . '" />
                        </a>
                      ';
            }

            echo '
                      </div>
                    </div>';
          }

          echo '
                </div>
              </div>

            </div>
            ';
        }
      }
      ?>

        </div>
        <div class="col s12 m12 l4 xl4">
            <ul class="tabs">
                <li class="tab col s6"><a class="active" href="#form_contacto" style="color: white;">Contactarme</a>
                </li>
                <li class="tab col s6"><a href="#contacto" style="color: white;">Datos de contacto</a></li>
            </ul>
            <div id="contacto" class="col s12 datos-contacto">
                <div class="card">
                    <div class="card-content datos">

                        <table class="highlight">
                            <tbody>
                                <tr>
                                    <td><span>Nombre: </span><?= $contacto['nombre'] ?></td>
                                </tr>
                                <tr>
                                    <td><span>Cargo: </span><?= $contacto['cargo'] ?></td>
                                </tr>
                                <tr>
                                    <td><span>Teléfono: </span><a
                                            href="https://api.whatsapp.com/send?phone=<?= $contacto['telefono'] ?>&text=Hola,%20me%20comunico%20desde%20la%20web%20384franquicias.com"
                                            target="_blank"><?= $contacto['telefono'] ?></a></td>
                                </tr>
                                <tr>
                                    <td><span>E-mail: </span><a href="mailto:<?= $contacto['mail'] ?>"
                                            target="_blank"><?= $contacto['mail'] ?></a></td>
                                </tr>
                                <tr>
                                    <td><span>Dirección: </span><?= $contacto['direccion'] ?></td>
                                </tr>
                                <tr>
                                    <td><span>Web: </span><a href="http://<?= $contacto['web'] ?>"
                                            target="_blank"><?= $contacto['web'] ?></a></td>
                                </tr>
                            </tbody>
                        </table>

                    </div>
                </div>
            </div>
            <div id="form_contacto" class="col s12">
                <div class="card">
                    <div class="card-content">
                        <div class="row">
                            <div class="col s12" style="text-align: center; margin: 20px 0">
                                <a href="https://api.whatsapp.com/send?phone=543515320234&text=Hola,%20me%20comunico%20desde%20la%20web%20384franquicias.com"
                                    target="_blank" style="text-decoration: none">
                                    <img src="<?= SERVER_URL ?>public/img/whatsapp-logo.png"
                                        style="width: 50px; height: auto">
                                    <p>Escribinos por Whatsapp!</p>
                                </a>
                            </div>
                        </div>
                        <div class="divider"></div>
                        <div class="row">
                            <div class="col s12" style="text-align: center; margin: 20px 0 10px">
                                <b>
                                    <p>o podés enviarnos un mail</p>
                                </b>
                            </div>
                        </div>
                        <div class="row">
                            <form id="formulario_contacto" action="" method="POST">
                                <div class="input-field col s6">
                                    <input id="nombre" type="text" class="validate required" name="nombre">
                                    <label for="nombre">Tu nombre y apellido (*)</label>
                                </div>
                                <div class="input-field col s6">
                                    <input id="email" type="email" class="validate required" name="email">
                                    <label for="email">Tu email (*)</label>
                                </div>
                                <div class="input-field col s6">
                                    <input id="telefono" type="text" class="validate required" name="telefono">
                                    <label for="telefono">Tu teléfono (*)</label>
                                </div>
                                <div class="input-field col s6">
                                    <select id="cod_pais" name="cod_pais" class="required">
                                        <option value="" disabled selected>Seleccionar</option>
                                        <?php
                    foreach ($paises as $key => $pais) {
                      echo '<option value="' . $pais['PaisCodigo'] . '">' . $pais['PaisNombre'] . '</option>';
                    }
                    ?>
                                    </select>
                                    <label>País (*)</label>
                                </div>
                                <div class="input-field col s6">
                                    <select id="id_provincia" name="id_provincia" class="required">
                                        <option value="" disabled selected>Seleccionar País</option>
                                    </select>
                                    <label>Provincia (*)</label>
                                </div>
                                <div class="input-field col s6">
                                    <input id="localidad" type="text" class="validate" name="localidad">
                                    <label for="localidad">Localidad</label>
                                </div>
                                <div class="input-field col s12">
                                    <select id="id_negocio" name="id_negocio" class="required">
                                        <?php
                    foreach ($negocios as $key => $negocio) {

                      $tipo = explode(':', $negocio['detalle']);

                      if ($tipo[0] == 'Franquicia') {
                        $tipo = 'franquicia';
                      } else {
                        $tipo = 'fondo_comercio';
                      }

                      echo '<option value="' . $tipo . ',' . $negocio['id'] . '">' . $negocio['detalle'] . '</option>';
                    }
                    ?>
                                    </select>
                                    <label>¿En qué estás interesado/a? (*)</label>
                                </div>
                                <div class="input-field col s12">
                                    <textarea id="mensaje" class="materialize-textarea required"
                                        name="mensaje"></textarea>
                                    <label for="mensaje">Tu mensaje (*)</label>
                                </div>
                                <div class="input-field col s12">
                                    <input type="hidden" name="id_marca" value="<?= $id_marca ?>">
                                    <button type="submit" class="waves-effect waves-light btn"
                                        style="width: 100%">ENVIAR</button>
                                </div>
                            </form>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="col s12 m12 l4 m4 right">

            <?php
      echo $publicidad;
      ?>

            <!--       <ul class="collapsible" style="margin-top: 20px">
        <li class="active">
          <div class="collapsible-header"><i class="material-icons">arrow_drop_down</i><b>TAMBIÉN TE PUEDE INTERESAR</b></div>
          <div class="collapsible-body">

            <div class="card franquicias">
              <div class="card-content">
                <div class="row valign-wrapper" style="padding: 10px">
                  <div style="width: 20%">
                    <img src="./assets/img/logo-384-chico.png" class="img-responsive" style="width: 100%">
                  </div>
                  <div>
                    <h5 style="margin-left: 30px"><b>Mc Donalds</b></h5>
                    <h6 style="margin-left: 30px"><b>Franquicia: Local</b></h6>
                  </div>
                </div>
              </div>
            </div>

            <div class="card franquicias">
              <div class="card-content">
                <div class="row valign-wrapper" style="padding: 10px">
                  <div style="width: 20%">
                    <img src="./assets/img/logo-384-chico.png" class="img-responsive" style="width: 100%">
                  </div>
                  <div>
                    <h5 style="margin-left: 30px"><b>Mc Donalds</b></h5>
                    <h6 style="margin-left: 30px"><b>Fondo de comercio: Stand</b></h6>
                  </div>
                </div>
              </div>
            </div>

          </div>
        </li>
      </ul> -->

        </div>
    </div>
</section>
<a href="https://api.whatsapp.com/send?phone=543518684146&text=Hola,%20me%20comunico%20desde%20la%20web%20384franquicias.com"
    target="_blank" style="text-decoration: none">
    <div class="row"
        style="position: fixed; bottom: 0; left: 0; background-color: #c8eaf1; margin: 20px; border-radius: 20px; padding: 20px; border: 2px solid #3d7986;">
        <div class="col s12" style="text-align: center;">
            <img src="<?= SERVER_URL ?>public/img/whatsapp-logo.png" style="width: 30px; height: auto">
            <p>¿Querés que te contemos más?</p>
        </div>
    </div>
</a>