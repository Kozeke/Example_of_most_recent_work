<?php

namespace App\Http\Controllers;

use App\Services\ActEditorPrintWord\ActEditorPrintWordService;
use App\Services\ActEditorSignaturesService;
use Illuminate\Contracts\Routing\ResponseFactory;
use Illuminate\Foundation\Application;
use Illuminate\Http\Response;
use PhpOffice\PhpWord\Exception\Exception;
use Symfony\Component\HttpFoundation\BinaryFileResponse;

/**
 * Class ActEditorDocumentPrintWordController
 * @package App\Http\Controllers
 *
 * @author Kozy-Korpesh Tolep
 */
class ActEditorDocumentPrintWordController extends Controller
{
    /**
     * @var ActEditorSignaturesService
     */
    private ActEditorSignaturesService $actEditorSignaturesService;

    /**
     * @var ActEditorPrintWordService
     */
    private ActEditorPrintWordService $actEditorPrintWordService;

    /**
     * ActEditorDocumentController constructor.
     * @param ActEditorSignaturesService $actEditorSignaturesService
     * @param ActEditorPrintWordService $actEditorPrintWordService
     */
    public function __construct(
        ActEditorSignaturesService $actEditorSignaturesService,
        ActEditorPrintWordService $actEditorPrintWordService
    ) {
        $this->actEditorPrintWordService = $actEditorPrintWordService;
        $this->actEditorSignaturesService = $actEditorSignaturesService;
    }

    /**
     * @param int $actEditorID
     * @return ResponseFactory|Application|Response|BinaryFileResponse
     */
    public function print(int $actEditorID)
    {
        try {
            $signaturesList = $this->actEditorSignaturesService->getList($actEditorID);
            return $this->actEditorPrintWordService->downloadWord($actEditorID, $signaturesList);
        } catch (Exception $e) {
            return $this->jsonException($e);
        }
    }
}