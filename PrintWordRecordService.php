<?php

namespace App\Services\ActEditorPrintWord;

use App\DTO\ActEditorPrint\ActEditorPrintWordRecordDTO;
use App\Extensions\ActEditor\PrintWord\ActEditorPrintWordTrait;
use PhpOffice\PhpWord\Element\Table;


/**
 * Class PrintWordRecords
 * @package App\Services\ActEditorPrintWord
 *
 * @author Tolep Kozy-Korpesh
 */
class PrintWordRecordService
{
    use ActEditorPrintWordTrait;

    /**
     * @var float
     */
    private const CONSTANT_K = 0.0169776;

    /**
     * @var string
     */
    private const DEFAULT_TEXT_VALUE = 'н/п';

    /**
     * @var int
     */
    private const TITLE_LAST_LINE_LENGTH = 0;

    /**
     * @var int
     */
    private const ADDITIONAL_WIDTH_TO_TABLE_IN_TWIP = 200;

    /**
     * @var int
     */
    private const ONE_LINE_FOR_TEXT_BREAK = 1;

    /**
     * @var ?int
     */
    private ?int $countSubCell = null;

    /**
     * @var array
     */
    private array $cellsOnOneLine = [];

    /**
     * @var bool
     */
    private bool $needSubCell = false;

    /**
     * @var array
     */
    private array $subCells = [];

    /**
     * @var string
     */
    private string $fieldTypeSizeOnPrint;

    /**
     * @var string|null
     */
    private ?string $titleTypeSizeOnPrint = null;

    /**
     * @var array
     */
    private array $subscriptRecordCells = [];

    /**
     * @var array
     */
    private array $oneLineCells = [];

    /**
     * @var int|null
     */
    private ?int $keyOfCurrentActEditorDocument;

    /**
     * @var array
     */
    private const PRINT_MAP = [
        'printOn' => true,
        'printOff' => false,
    ];

    /**
     * @var array
     */
    private const PROPS_FOR_TEMPLATE = [
        'qualityDocName' => 'name',
        'qualityDocNum' => 'number',
        'qualityDocDate' => 'date',
        'projectDocumentation' => 'name',
        'projectDocumentationNumber' => 'number',
        'projectDocumentationDate' => 'date',
    ];

    /**
     * @param ActEditorPrintWordRecordDTO $actEditorPrintWordRecordDTO
     * @param Table|null $table
     * @return void
     */
    public function ifTabTypeRecord(
        ActEditorPrintWordRecordDTO $actEditorPrintWordRecordDTO,
        ?Table &$table = null
    ): void {
        $this->massAssignmentOfPropertiesFromDTO($actEditorPrintWordRecordDTO);
        $this->table = &$table;
        $this->isRecordType = true;
        $actEditorDocument = $this->actEditorDocuments[$this->keyOfCurrentActEditorDocument];
        $this->fontSize = $actEditorDocument['fontSize'] ?? (int)$this->predefinedStyles['fontSize'];
        //не добавляем новую строку если предыдущая запись имеет тот же lineNumberOnPrint
        if ($this->keyOfCurrentActEditorDocument === 0 || $this->isActEditorDocumentHasSameLineNumber(
                $actEditorDocument,
                $this->keyOfCurrentActEditorDocument - 1
            )) {
            $this->countSubCell = null;
            $this->section->addTextBreak(
                self::ONE_LINE_FOR_TEXT_BREAK,
                $this->predefinedStyles['sizeBreak'],
                $this->predefinedStyles['spaceBreak']
            );
            $this->table = $this->section->addTable();
            $this->table->addRow();
        }
        $this->countSubCell = is_null($this->countSubCell) ? 0 : $this->countSubCell + 1;
        $isNextActEditorDocSameLineNumber = $this->isActEditorDocumentHasSameLineNumber($actEditorDocument, $this->keyOfCurrentActEditorDocument + 1);
        if (
            $this->isSubCellNecessary($actEditorDocument)
        ) {
            $this->countSubCell++;
        }
        $this->printRecord($actEditorDocument);
        $this->cellsOnOneLine[] = $actEditorDocument;
        if (
            ($this->isCurrentActEditorDocumentIsLast() || !$isNextActEditorDocSameLineNumber)
            && $this->needSubCell
        ) {
            $this->setCountSubCell();
            $this->createSubscriptCells();
        }
        $this->cellsOnOneLine = $isNextActEditorDocSameLineNumber ? $this->cellsOnOneLine : [];
        $this->isRecordType = false;
    }

    /**
     * @param ActEditorPrintWordRecordDTO $actEditorPrintWordRecordDTO
     * @return void
     */
    public function ifTabTypeSignature(ActEditorPrintWordRecordDTO $actEditorPrintWordRecordDTO): void
    {
        $this->massAssignmentOfPropertiesFromDTO($actEditorPrintWordRecordDTO);
        $this->isSignatureType = true;
        $this->countSubCell = 1;
        $this->table = $this->section->addTable();
        $this->table->addRow();
        $this->printRecord($this->currentActEditorDocument);
        $this->cellsOnOneLine[] = $this->currentActEditorDocument;
        $this->section->addTextBreak(
            self::ONE_LINE_FOR_TEXT_BREAK,
            $this->predefinedStyles['sizeBreak'],
            $this->predefinedStyles['spaceBreak']
        );
        if ($this->needSubCell) {
            $this->createSubscriptCells();
        }
        $this->cellsOnOneLine = [];
        $this->isSignatureType = false;
    }

    /**
     * @param array $actEditorDocument
     * @param $key
     * @return bool
     */
    private function isActEditorDocumentHasSameLineNumber(array $actEditorDocument, $key): bool
    {
        return ($this->actEditorDocuments[$key]['lineNumberOnPrint'] ?? '') === $actEditorDocument['lineNumberOnPrint'];
    }

    /**
     * @return bool
     */
    private function isCurrentActEditorDocumentIsLast(): bool
    {
        return count($this->actEditorDocuments) - 1 === $this->keyOfCurrentActEditorDocument;
    }

    /**
     * @return void
     */
    private function setCountSubCell(): void
    {
        $this->countSubCell = !!implode(
            ', ',
            array_filter(array_map(fn($item) => $item['subsText'], $this->subCells))
        ) ? $this->countSubCell : 0;
    }

    /**
     * @param array $actEditorDocument
     * @return bool
     */
    private function isSubCellNecessary(array $actEditorDocument): bool
    {
        return $actEditorDocument['viewTitleOnPrint']
            && $actEditorDocument['fieldTypeSizeOnPrint'] !== 'string'
            && $actEditorDocument['subscriptOnPrint']
            && $this->countSubCell === 0;
    }

}