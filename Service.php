<?php

namespace App\Services\ActEditorPrintWord;

use App\DTO\ActEditorPrint\ActEditorPrintWordRecordDTO;
use App\DTO\ActEditorPrint\ActEditorPrintWordTableDTO;
use App\Extensions\ActEditor\PrintWord\ActEditorPrintWordTrait;
use App\Models\ActEditor\ActEditor;
use App\Models\ActEditor\ActEditorDocument;
use App\Models\Entrance\Act\EntranceJournalAct;
use App\Models\Journal\JournalExecutiveDocumentation;
use Illuminate\Database\Eloquent\Collection;
use Illuminate\Support\Collection as SupportCollection;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\Language;
use Symfony\Component\HttpFoundation\BinaryFileResponse;

/**
 * Class ActEditorPrintWordService
 * @package App\Services\ActEditorPrintWord
 *
 * @author Kozy-Korpesh Tolep
 */
class ActEditorPrintWordService
{
    use ActEditorPrintWordTrait;

    /**
     * @var string
     */
    private const TABLE_STYLE_NAME_FOR_ACTS = 'Раздел 1';

    /**
     * @var string
     */
    private const TABLE_STYLE_NAME_FOR_SIGNS = 'Раздел 2';

    /**
     * @var array
     */
    private array $supplementsByDocumentIdMap = [];

    /**
     * @var int
     */
    private const DEFAULT_SPACING_TWIP = 240;

    /**
     * @param int $actEditorID
     * @param Collection $signaturesList
     * @return BinaryFileResponse
     * @throws Exception
     */
    public function downloadWord(
        int $actEditorID,
        Collection $signaturesList
    ): BinaryFileResponse {
        if (isset($this->model, $this->model->act, $this->model->act->name)) {
            $this->actEditorNameToLower = mb_strtolower($this->model->act->name, 'UTF-8');
        }
        $this->isEntranceJournalAct = $this->model instanceof EntranceJournalAct;
        $isJournalExecutiveDocumentation = $this->model instanceof JournalExecutiveDocumentation;
        $actEditorDocuments = $this->getActEditorDocuments($actEditorID);
        $docStyleSettings = $this->getDocParams($actEditorID);
        $this->setPredefinedStyles($docStyleSettings);

        $phpWord = new PhpWord();
        $this->setDefaultSettings($phpWord, $docStyleSettings);
        $actEditorDocuments = $this->mergeAndOrderActEditorDocumentsWithSignatures(
            $actEditorDocuments,
            $signaturesList
        );

        $this->section = $phpWord->addSection($this->predefinedStyles["styleSection"]);
        $phpWord = $this->printActEditorDocumentRecords($actEditorDocuments, $phpWord);

        //print signatures for JournalExecutiveDocumentation
        if ($isJournalExecutiveDocumentation) {
            $phpWord = $this->printTableWithActSignatures($phpWord);
        }
        $pathToTempFile = $this->createAndSaveFile($phpWord, $docStyleSettings->name);
        return response()->download($pathToTempFile)->deleteFileAfterSend();
    }


    /**
     * @param PhpWord $phpWord
     * @param ActEditor $docStyleSettings
     * @return void
     */
    private function setDefaultSettings(phpWord $phpWord, ActEditor $docStyleSettings)
    {
        $phpWord->addParagraphStyle('StyleSubs', array('align' => 'center', 'spaceAfter' => 10));
        $phpWord->getSettings()->setThemeFontLang(new Language(Language::RU_RU));
        $phpWord->setDefaultFontName($docStyleSettings->fontFamily);
        $phpWord->setDefaultFontSize($docStyleSettings->fontSize);
    }

    /**
     * @param PhpWord $phpWord
     * @param string $fileName
     * @return string
     * @throws Exception
     */
    private function createAndSaveFile(phpWord $phpWord, string $fileName): string
    {
        $objWriter = IOFactory::createWriter($phpWord);
        $pathToTempFile = Storage::path('public/' . $fileName . '.docx');
        $objWriter->save($pathToTempFile);
        return $pathToTempFile;
    }


    /**
     * @param ActEditor $docStyleSettings
     * @return void
     */
    private function setPredefinedStyles(ActEditor $docStyleSettings): void
    {
        $this->predefinedStyles['styleSection']['marginLeft'] = Converter::cmToTwip($docStyleSettings->leftIndent);
        $this->predefinedStyles['styleSection']['marginRight'] = Converter::cmToTwip($docStyleSettings->rightIndent);
        $this->predefinedStyles['styleSection']['marginTop'] = Converter::cmToTwip($docStyleSettings->topIndent);
        $this->predefinedStyles['styleSection']['marginBottom'] = Converter::cmToTwip($docStyleSettings->bottomIndent);
        $spaceSizes = ['1' => 1, '2' => self::DEFAULT_SPACING_TWIP];
        ['lineSpacing' => $lineSpacing] = $docStyleSettings;
        $this->predefinedStyles['spacing'] =
            $spaceSizes[$lineSpacing]
            ?? (float)$lineSpacing * self::DEFAULT_SPACING_TWIP - self::DEFAULT_SPACING_TWIP;
        $this->predefinedStyles['fontSize'] = $docStyleSettings->fontSize;
        $this->predefinedStyles['styleComments']['size'] = $docStyleSettings->fontSize - 3;
        $this->isDatesCheck = $docStyleSettings->isDatesCheck;
    }

    /**
     * @return void
     */
    private function setSupplementsByDocumentIdMap(): void
    {
        if ($this->model instanceof JournalExecutiveDocumentation) {
            $supplementRows = $this->getSupplementRows();
            $supplementRows->each(
                function ($supplementRow) {
                    $supplementArray = $this->supplementsByDocumentIdMap[$supplementRow->document_id] ?? null;
                    $this->supplementsByDocumentIdMap[$supplementRow->document_id] = isset($supplementArray)
                        ? [...$supplementArray, $supplementRow]
                        : [$supplementRow];
                }
            );
        }
    }

    /**
     * @return void
     */
    private function calculatePageWidthInTwipWithoutMargins(): void
    {
        $pageWidthInTwip = $this->section->getStyle()->getPageSizeW();
        $marginLeftInTwip = $this->section->getStyle()->getMarginLeft();
        $marginRightInTwip = $this->section->getStyle()->getMarginRight();
        $this->pageWidthInTwipWithoutMargins = $pageWidthInTwip - $marginLeftInTwip - $marginRightInTwip;
    }

    /**
     * @param SupportCollection $actEditorDocuments
     * @param PhpWord $phpWord
     * @return PhpWord
     */
    private function printActEditorDocumentRecords(
        SupportCollection $actEditorDocuments,
        PhpWord $phpWord
    ): PhpWord {
        $this->section->addHeader()->addPreserveText('{PAGE}', [], ["alignment" => Jc::END]);
        $this->calculatePageWidthInTwipWithoutMargins();
        $this->setSupplementsByDocumentIdMap();

        foreach ($actEditorDocuments as $key => $actEditorDocument) {
            $this->actEditorDocuments = $actEditorDocuments;
            ['tab_type' => $tabType] = $actEditorDocument;
            if ($actEditorDocument['viewTitleOnPrint'] || $actEditorDocument['viewFieldOnPrint']) {
                if ($tabType === 'table' || $tabType === 'supplement') {
                    $this->replaceSupplementTabTypeOnTable($tabType, $actEditorDocument);
                    $phpWord->addTableStyle(self::TABLE_STYLE_NAME_FOR_ACTS, $this->predefinedStyles["styleTable"], []);
                    (new PrintWordTableService())->printTable(
                        ActEditorPrintWordTableDTO::createFromArray(
                            [
                                "section" => $this->section,
                                "page_width_in_twip_without_margins" => $this->pageWidthInTwipWithoutMargins,
                                "current_act_editor_document" => $actEditorDocument,
                                "model" => $this->model,
                                "object" => $this->object
                            ]
                        )
                    );
                } elseif ($tabType === 'record') {
                    (new PrintWordRecordService())->ifTabTypeRecord(
                        ActEditorPrintWordRecordDTO::createFromArray(
                            [
                                "key_of_current_act_editor_document" => $key,
                                "section" => $this->section,
                                "page_width_in_twip_without_margins" => $this->pageWidthInTwipWithoutMargins,
                                "act_editor_documents" => $actEditorDocuments,
                                "model" => $this->model,
                                "object" => $this->object
                            ]
                        ),
                        $table
                    );
                } elseif ($tabType === 'signature') {
                    (new PrintWordRecordService())->ifTabTypeSignature(
                        ActEditorPrintWordRecordDTO::createFromArray(
                            [
                                "section" => $this->section,
                                "page_width_in_twip_without_margins" => $this->pageWidthInTwipWithoutMargins,
                                "current_act_editor_document" => $actEditorDocument,
                                "model" => $this->model,
                                "object" => $this->object
                            ]
                        )
                    );
                }
            }
        }
        return $phpWord;
    }

    /**
     * @param string $tabType
     * @param array $actEditorDocument
     */
    private function replaceSupplementTabTypeOnTable(string $tabType, array &$actEditorDocument): void
    {
        if ($tabType === 'supplement') {
            $actEditorDocument['isMarginTop'] = true;
            $actEditorDocument['tab_type'] = 'table';
            $actEditorDocument['isSupplement'] = true;
            $actEditorDocument['titleTypeSizeOnPrint'] = 'left';
            $actEditorDocument['titleHorizontalAlignmentOnPrint'] = 'newlineInterlinearEnum';
        }
    }

    /**
     * @param PhpWord $phpWord
     * @return PhpWord
     */
    private function printTableWithActSignatures(PhpWord $phpWord): PhpWord
    {
        $phpWord->addTableStyle(self::TABLE_STYLE_NAME_FOR_SIGNS, $this->predefinedStyles["styleTableForSigns"],);
        (new PrintWordTableService())->printTableWithActSignatures(
            ActEditorPrintWordTableDTO::createFromArray(
                [
                    "section" => $this->section,
                    "page_width_in_twip_without_margins" => $this->pageWidthInTwipWithoutMargins,
                    "signatures_with_kcp" => $this->signaturesWithKcp,
                    "signatures" => $this->signatures
                ]
            )
        );
        return $phpWord;
    }
}
