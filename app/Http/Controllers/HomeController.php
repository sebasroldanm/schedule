<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\TemplateProcessor;

class HomeController extends Controller
{
    /**
     * Create a new controller instance.
     *
     * @return void
     */
    public function __construct()
    {
        $this->middleware('auth');
    }

    /**
     * Show the application dashboard.
     *
     * @return \Illuminate\Http\Response
     */
    public function index()
    {

        // $documento = new PhpWord();
        // $seccion = $documento->addSection();
        // $seccion->addText(
        //     htmlspecialchars(
        //     'Primer texto - Texto sin formato'
        //     )
        // );
        // $seccion->addText(
        //     htmlspecialchars(
        //     'Segundo texto con formato'
        //     ),
        //     array('name' => 'Arial', 'size' => '12', 'bold' => 'true')
        // );
        // $seccion->addTextBreak(1);
        // $fuente = new Font();
        // $fuente->setBold(true);
        // $fuente->setName('Tahoma');
        // $fuente->setSize(16);
        // $fuente->setColor('9F81F7');
        // $texto = $seccion->addText(htmlspecialchars(
        //     'Cuarto texto con formato'
        // ));
        // $texto->setFontStyle($fuente);
        // $estilo_tabla = array(
        //     'borderColor' => 'F2F2F2',
        //     'borderSize' => '5',
        //     'cellMargin' => '20',
        //     'bgColor' => '088A68',
        //   );
        // $primera_fila = array('bgColor' => 'F2F2F2');
        // $documento->addTableStyle('mitabla',$estilo_tabla, $primera_fila);
        // $tabla = $seccion->addTable('mitabla');
        // for ($row = 1; $row <= 8; $row++) {
        //     $tabla->addRow();
        //     for ($cell = 1; $cell <= 3; $cell++) {
        //         if($row ==1)
        //             $tabla->addCell(200)->addText(htmlspecialchars('primera'));
        //         else
        //             $tabla->addCell(200)->addText(htmlspecialchars('celda'));
        //     }
        // }
        // $objWriter = IOFactory::createWriter($documento, 'Word2007');
        // $objWriter->save('Documento1.docx');

        $templateWord = new TemplateProcessor('AutoEvaluacionInforme2015sistemasfacatativaV2.docx');

        $nombre_universidad = "Universidad de Cundinamarca";
        $direccion_universidad = "Av Cajica";
        $snies = "2120 Fusagasuga   2234 Facatativa     6577 Chia";
        $norma_creacion = "Ordenanza número 045 del 19 de diciembre de 1969";
        $estudiantes_matriculados = "14566";
        $metodologia = "Presencial";
        $p_p = "39";
        $p_tm = "490";
        $p_mt = "80";
        $p_c = "430";
        $graduados = "20987";
        $nombre_tabla_institucion = "Fuente Boletín estadístico, datos actualizados a junio de 2019";
        $mision = "La Universidad de Cundinamarca es una institución pública local y translocal del Siglo XXI, caracterizada por ser una organización social de conocimiento, democrática, autónoma, formadora, agente de la transmodernidad que incorpora los consensos mundiales de la humanidad y las buenas prácticas de gobernanza universitaria, cuya calidad se genera desde los procesos de enseñanza-aprendizaje, investigación e innovación, interacción universitaria, internacionalización y bienestar universitario";
        $vision = "La Universidad de Cundinamarca será reconocida por la sociedad, en el ámbito local, regional, nacional e internacional, como generadora de conocimiento relevante y pertinente, centrada en el cuidado de la vida, la naturaleza, el ambiente, la humanidad y la convivencia";

        $templateWord->setValue('nombre_universidad', $nombre_universidad);
        $templateWord->setValue('direccion_universidad', $direccion_universidad);
        $templateWord->setValue('snies', $snies);
        $templateWord->setValue('norma_creacion', $norma_creacion);
        $templateWord->setValue('estudiantes_matriculados', $estudiantes_matriculados);
        $templateWord->setValue('metodologia', $metodologia);
        $templateWord->setValue('p_p', $p_p);
        $templateWord->setValue('p_tm', $p_tm);
        $templateWord->setValue('p_mt', $p_mt);
        $templateWord->setValue('p_c', $p_c);
        $templateWord->setValue('graduados', $graduados);
        $templateWord->setValue('nombre_tabla_institucion', $nombre_tabla_institucion);
        $templateWord->setValue('mision_universidad', $mision);
        $templateWord->setValue('vision_universidad', $vision);


        $templateWord->saveAs('InformeAuto.docx');

        $pathToFile = public_path(). "\InformeAuto.docx";
// dd($pathToFile);
        return response()->download($pathToFile);

        return view('home');
    }
}
