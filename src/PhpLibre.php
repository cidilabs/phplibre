<?php

namespace CidiLabs\PhpLibre;

use DOMDocument;
use DOMXPath;
use Ramsey\Uuid\Uuid;

class PhpLibre
{

    private $bin;

    private $outputDir;

    // extensions and filters for LibreOffice
    // https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html
    private $exportFilters = [
        'doc' => [
            'html' => 'html:HTML:EmbedImages'
        ],
        'DOC' => [
            'html' => 'html:HTML:EmbedImages'
        ],
        'docx' => [
            'html' => 'html:HTML:EmbedImages'
        ],
        'DOCX' => [
            'html' => 'html:HTML:EmbedImages'
        ],
        'pdf' => [
            'html' => 'html:XHTML Impress File'
        ],
        'pptx' =>  [
            'html' => 'html:XHTML Impress File'
        ]
    ];
    private $infilterOptions = [
        'pdf' => 'impress_pdf_import',
    ];

    // Response object that gets returned to Udoit
    private $responseObject = [
        'data' => [
            'taskId' => '',
            'filePath' => '',
            'relatedFiles' => [],
            'status' => ''
        ],
        'errors' => []
    ];

    private $sOfficePort = null;
    private $tempDir;

    public function __construct($bin = 'soffice', $outputDir = 'alternates')
    {
        $this->bin = $bin;
        $this->outputDir = $outputDir;
        $this->tempDir = sys_get_temp_dir();
    }

    public function supports()
    {
        return [
            'input' => ['pdf', 'doc','pptx'],
            'output' => ['html']
        ];
    }

    public function convertFile($options)
    {
        $fileUrl = $options['fileUrl'];
        $extension = $options['fileType'];
        $fileName = $options['fileName'];
        $format = $options['format'];
        $taskId = Uuid::uuid4()->toString();
        $newFilename = $taskId . '.' . $format;
        $supportedExtensions = $this->getAllowedConverter($extension);

        //Check for valid input file extension
        if (!array_key_exists($extension, $this->getAllowedConverter())) {
            $this->responseObject['errors'][] = "Input file extension not supported -- " . $extension;
        }

        if (!in_array($format, $supportedExtensions)) {
            $this->responseObject['errors'][] = "Output extension({$format}) not supported for input file({$fileUrl})";
        }

        if (!file_put_contents($fileName, file_get_contents($fileUrl))) {
            $this->responseObject['errors'][] = "File downloading failed.";
        }

        if (!empty($this->responseObject['errors'])) {
            return $this->responseObject;
        }

        if (!is_dir($this->outputDir)) {
            mkdir($this->outputDir, 0777, true);
        }

        $shell = $this->exec($this->makeCommand($extension, $format, $fileName));

        if (0 != $shell['return']) {
            $this->responseObject['errors'][] = "Conversion Failure! Contact your institution's UDOIT admin. Error: " . $shell['return'];
            return $this->responseObject;
        }

        $DS = DIRECTORY_SEPARATOR;
        $outdir = $this->outputDir;
        $tmpName = pathinfo($fileName, PATHINFO_FILENAME) . '.' . $format;

        rename($outdir . $DS . $tmpName, $outdir . $DS . $newFilename);

        $this->responseObject['data']['taskId'] = $taskId;

        return $this->responseObject;
    }

    public function convertToHtml($filepath)
    {
        $extension = pathinfo($filepath, PATHINFO_EXTENSION);

        //Check for valid input file extension
        if (!array_key_exists($extension, $this->getAllowedConverter())) {
            throw new \Exception("Input file extension not supported -- " . $extension);
        }

        if (!is_dir($this->outputDir)) {
            mkdir($this->outputDir, 0777, true);
        }

        $shell = $this->exec($this->makeCommand($extension, 'html', $filepath));

        if (0 != $shell['return']) {
            throw new \Exception("Conversion Failure! Contact your institution's UDOIT admin. Error: " . $shell['return'] . ": output :" . $shell['stdout']. ": error :" . $shell['stderr']);
        }

        $htmlFilepath = str_replace($extension, 'html', $filepath);

        if (!file_exists($htmlFilepath)) {
            return null;
        }

        $out = file_get_contents($htmlFilepath);

        unlink($htmlFilepath);

        return $out;
    }

    public function isReady($taskId)
    {
        $result = glob($this->outputDir . '/' . $taskId . '.*');

        return !empty($result);
    }

    public function getFileUrl($taskId, $options = [])
    {
        $result = glob($this->outputDir . '/' . $taskId . '.*');

        // TODO: handle related files as well

        if (!empty($result)) {
            $this->responseObject['data']['filePath'] = $result[0];
        } else {
            $this->responseObject['errors'][] = "No file found for taskId: " . $taskId;
        }

        return $this->responseObject;
    }

    public function deleteFile($fileUrl)
    {
        if (file_exists($fileUrl)) {
            unlink($fileUrl);
        } else {
            $this->responseObject['errors'][] = "File not found";
        }

        return $this->responseObject;
    }

    /**
     * Helpers
     **/

    protected function makeCommand($inputExtension, $outputExtension, $filename)
    {
        $oriFile = escapeshellarg($filename);
        $dirname = $this->outputDir;

        //Finds an output filter that corresponds to the input and output types
        $outputExtension = !empty($this->exportFilters[$inputExtension][$outputExtension]) ? $this->exportFilters[$inputExtension][$outputExtension] : $outputExtension;

        //Determines the infilter based on the input type
        $infilterArg = !empty($this->infilterOptions[$inputExtension]) ? $this->infilterOptions[$inputExtension] : '';
        $infilter = "--infilter=\"" . $infilterArg . "\" ";

        while (is_null($this->sOfficePort)){
            $tmpPort = rand(8100,8999);
            if (!is_dir("{$this->tempDir}/SOffice_Process{$this->sOfficePort}")) {
                $this->sOfficePort = $tmpPort;
            }
        }

        return "{$this->bin} --headless --headless --norestore --nolockcheck  -env:SingleAppInstance=\"false\" -env:UserInstallation=\"file:///{$this->tempDir}/SOffice_Process{$this->sOfficePort}\" --accept=\"socket,host=localhost,port={$this->sOfficePort};urp;\"  " . $infilter . "--convert-to \"{$outputExtension}\" {$oriFile} --outdir {$dirname}";
    }

    protected function open($filename)
    {
        if (!file_exists($filename) || false === realpath($filename)) {
            print('File does not exist --' . $filename);
            return false;
        }

        return true;
    }

    private function getAllowedConverter($extension = null)
    {
        $allowedConverter = [
            '' => ['pdf'],
            'pptx' => ['pdf'],
            'ppt' => ['pdf'],
            'pdf' => ['pdf', 'html'],
            'docx' => ['pdf', 'odt', 'html'],
            'doc' => ['pdf', 'odt', 'html'],
            'wps' => ['pdf', 'odt', 'html'],
            'dotx' => ['pdf', 'odt', 'html'],
            'docm' => ['pdf', 'odt', 'html'],
            'dotm' => ['pdf', 'odt', 'html'],
            'dot' => ['pdf', 'odt', 'html'],
            'odt' => ['pdf', 'html'],
            'xlsx' => ['pdf'],
            'xls' => ['pdf'],
            'png' => ['pdf'],
            'jpg' => ['pdf'],
            'jpeg' => ['pdf'],
            'jfif' => ['pdf'],
            'PPTX' => ['pdf'],
            'PPT' => ['pdf'],
            'PDF' => ['pdf', 'html'],
            'DOCX' => ['pdf', 'odt', 'html'],
            'DOC' => ['pdf', 'odt', 'html'],
            'WPS' => ['pdf', 'odt', 'html'],
            'DOTX' => ['pdf', 'odt', 'html'],
            'DOCM' => ['pdf', 'odt', 'html'],
            'DOTM' => ['pdf', 'odt', 'html'],
            'DOT' => ['pdf', 'odt', 'html'],
            'ODT' => ['pdf', 'html'],
            'XLSX' => ['pdf'],
            'XLS' => ['pdf'],
            'PNG' => ['pdf'],
            'JPG' => ['pdf'],
            'JPEG' => ['pdf'],
            'JFIF' => ['pdf'],
            'Pptx' => ['pdf'],
            'Ppt' => ['pdf'],
            'Pdf' => ['pdf'],
            'Docx' => ['pdf', 'odt', 'html'],
            'Doc' => ['pdf', 'odt', 'html'],
            'Wps' => ['pdf', 'odt', 'html'],
            'Dotx' => ['pdf', 'odt', 'html'],
            'Docm' => ['pdf', 'odt', 'html'],
            'Dotm' => ['pdf', 'odt', 'html'],
            'Dot' => ['pdf', 'odt', 'html'],
            'Ddt' => ['pdf', 'html'],
            'Xlsx' => ['pdf'],
            'Xls' => ['pdf'],
            'Png' => ['pdf'],
            'Jpg' => ['pdf'],
            'Jpeg' => ['pdf'],
            'Jfif' => ['pdf'],
            'rtf'  => ['docx', 'txt', 'pdf'],
            'txt'  => ['pdf', 'odt', 'doc', 'docx', 'html'],
        ];

        if (null !== $extension) {
            if (isset($allowedConverter[$extension])) {
                return $allowedConverter[$extension];
            }

            return [];
        }

        return $allowedConverter;
    }

    /**
     * More intelligent interface to system calls.
     *
     * @see http://php.net/manual/en/function.system.php
     *
     * @param string $cmd
     * @param string $input
     *
     * @return array
     */
    private function exec($cmd, $input = '')
    {
        $process = proc_open($cmd, [0 => ['pipe', 'r'], 1 => ['pipe', 'w'], 2 => ['pipe', 'w']], $pipes);
        if (false === $process) {
            print('Cannot obtain ressource for process to convert file');
        }

        fwrite($pipes[0], $input);
        fclose($pipes[0]);
        $stdout = stream_get_contents($pipes[1]);
        fclose($pipes[1]);
        $stderr = stream_get_contents($pipes[2]);
        fclose($pipes[2]);
        $rtn = proc_close($process);

        $this->deleteDirectory("{$this->tempDir}/SOffice_Process{$this->sOfficePort}");

        return [
            'stdout' => $stdout,
            'stderr' => $stderr,
            'return' => $rtn,
        ];
    }

    private function deleteDirectory($dir) {
        if (!file_exists($dir)) {
            return true;
        }
    
        if (!is_dir($dir)) {
            return unlink($dir);
        }
    
        foreach (scandir($dir) as $item) {
            if ($item == '.' || $item == '..') {
                continue;
            }
    
            if (!$this->deleteDirectory($dir . DIRECTORY_SEPARATOR . $item)) {
                return false;
            }
    
        }
    
        return rmdir($dir);
    }
}
