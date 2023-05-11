<?php

namespace App\Services\ActEditorPrintWord;

use App\DTO\ActEditorPrint\ActEditorPrintWordTableDTO;
use App\Extensions\ActEditor\PrintWord\ActEditorPrintWordTrait;
use Illuminate\Support\Collection as SupportCollection;
use PhpOffice\PhpWord\Shared\Html;


/**
 * Class ActEditorPrintWordService
 * @package App\Services
 *
 * @author Kozy-Korpesh Tolep
 */
class PrintWordTableService
{
    use ActEditorPrintWordTrait;

    /**
     * @var int
     */
    private const ONE_LINE_FOR_TEXT_BREAK = 1;

    /**
     * @var int
     */
    private const FIVE_LINES_FOR_TEXT_BREAK = 5;


    /**
     * @var string
     */
    private const TABLE_STYLE_NAME_FOR_ACTS = 'Раздел 1';

    /**
     * @var string
     */
    private const TABLE_STYLE_NAME_FOR_SIGNS = 'Раздел 2';

    /**
     * @var string
     */
    private const STYLE_CELL_WIDTH = "width:50%";

    /**
     * @var array
     */
    private array $headerValuesForSignsWithKCPTable = [
        "Владелец сертификата: организация, сотрудник",
        "Сертификат: серийный номер, период действия",
        "Дата и время подписания"
    ];

    /**
     * @var array
     */
    private array $headerValuesForSignsTable = [
        "Владелец сертификата: организация, сотрудник",
        "Дата и время подписания"
    ];

    /**
     * @var string
     */
    private string $mainHeaderOfSignsTable = "Документ подписан и передан через веб-систему Adept";

    /**
     * @var string
     */
    private string $signConfirmedText = "Подпись соответствует файлу документа";

    /**
     * @var string
     */
    private string $borderThicknessAndColor = "1px solid black";

    /**
     * @var array
     */
    private array $supplementsByDocumentIdMap = [];

    /**
     * @var string
     */
    private string $orientation = 'portrait';

    private ?array $columnTypeTextMap = null;
    private ?array $columnTypeDatesMap = null;
    private ?array $columnTypePartnersMap = null;


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
     * @var int
     */
    private const DEFAULT_SPACING_TWIP = 240;

    /**
     * @var int
     */
    private const GRID_SPAN_COUNT = 2;

    /**
     * @param ActEditorPrintWordTableDTO $actEditorPrintWordTableDTO
     * @return void
     */
    public function printTable(ActEditorPrintWordTableDTO $actEditorPrintWordTableDTO): void
    {
        $this->massAssignmentOfPropertiesFromDTO($actEditorPrintWordTableDTO);
        if ($this->currentActEditorDocument['print'] === 'table') {
            $this->printAsTable($this->currentActEditorDocument);
        } else {
            $this->printAsString($this->currentActEditorDocument);
        }
        $this->section->addTextBreak(
            self::ONE_LINE_FOR_TEXT_BREAK,
            $this->predefinedStyles['sizeBreak'],
            $this->predefinedStyles['spaceBreak']
        );
    }

    /**
     * @param array $actEditorDocument
     * @return void
     */
    private function printAsTable(array $actEditorDocument): void
    {
        ['columns' => $columns, 'is_custom_numeration' => $isCustomNumeration] = $actEditorDocument;
        $filteredColumns = array_filter(
            $columns,
            fn($column) => $column['print'] === 'printOn'
        );
        if (count($filteredColumns) !== 0) {
            $this->section->addText(
                htmlspecialchars($actEditorDocument['fieldName']),
                [
                    'size' => $actEditorDocument['fontSize'],
                    'bold' => $actEditorDocument['titleBoldOnPrint'],
                    'italic' => $actEditorDocument['titleItalicOnPrint'],
                ],
                [
                    'align' => $actEditorDocument['titleHorizontalAlignmentOnPrint'],
                    'spaceAfter' => 0,
                ]
            );
            $this->table = $this->section->addTable(self::TABLE_STYLE_NAME_FOR_ACTS);
            $this->table->addRow();
            $filteredColumns = $this->getOrderedByNumerationColumns($filteredColumns, $isCustomNumeration);
            $this->printHeader($filteredColumns, $actEditorDocument);
            $this->printNumerationRow($filteredColumns, $actEditorDocument, $isCustomNumeration);
            $this->printBody($filteredColumns, $actEditorDocument);
        }
    }

    /**
     * @param array $columns
     * @param array $actEditorDocument
     * @param string|null $keyForValue
     * @return void
     */
    private function printRow(
        array $columns,
        array $actEditorDocument,
        ?string $keyForValue = 'field_name'
    ): void {
        foreach ($columns as $column) {
            $columnSize = $this->calculateLengthOfTextInTwip(
                'procent',
                $column['field_procent_size_on_screen'],
                $column['field_name']
            );
            $this->table->addCell($columnSize, $this->predefinedStyles['styleCell'])->addText(
                htmlspecialchars($keyForValue ? $column[$keyForValue] : ''),
                ['size' => $actEditorDocument['fontSize']],
                ['align' => 'center', 'spaceAfter' => 0, 'spaceBefore' => 0]
            );
        }
    }

    /**
     * @param SupportCollection $rows
     * @param array $columns
     * @param array $actEditorDocument
     * @return void
     */
    private function printEmptyRow(SupportCollection $rows, array $columns, array $actEditorDocument): void
    {
        if ($rows->count() === 0) {
            $this->table->addRow();
            $this->printRow($columns, $actEditorDocument, null);
        }
    }

    /**
     * @param array $columns
     * @param array $actEditorDocument
     * @return void
     */
    private function printHeader(array $columns, array $actEditorDocument): void
    {
        $this->printRow($columns, $actEditorDocument);
    }

    /**
     * @param array $columns
     * @param array $actEditorDocument
     * @param bool $isCustomNumeration
     * @return void
     */
    private function printNumerationRow(array $columns, array $actEditorDocument, bool $isCustomNumeration): void
    {
        if ($isCustomNumeration) {
            $this->table->addRow();
            $this->printRow($columns, $actEditorDocument, 'custom_column_number');
        }
    }

    /**
     * @param array $columns
     * @param array $actEditorDocument
     * @return void
     */
    private function printBody(array $columns, array $actEditorDocument): void
    {
        $rows = $this->getRows($columns);
        $rows->each(function ($item) use ($actEditorDocument) {
            $this->table->addRow();
            $this->printRow($item->get('columns'), $actEditorDocument, 'valueForTable');
        });
        $this->printEmptyRow($rows, $columns, $actEditorDocument);
    }

    /**
     * @param array $columns
     * @param bool $isCustomNumeration
     * @return array
     */
    private function getOrderedByNumerationColumns(array $columns, bool $isCustomNumeration): array
    {
        if ($isCustomNumeration) {
            $key = 'custom_column_number';

            usort(
                $columns,
                function ($columnA, $columnB) use ($key) {
                    if ($columnA[$key] === null && $columnB[$key] !== null) {
                        return 1;
                    }
                    if ($columnA[$key] !== null && $columnB[$key] === null) {
                        return -1;
                    }
                    return $columnA[$key] - $columnB[$key];
                }
            );

            if ($columns[0][$key] === null) {
                $columns = array_map(
                    function ($column) use ($key) {
                        $column[$key] = intval($column['name']);
                        return $column;
                    },
                    $columns
                );
            }
        }

        return $columns;
    }

    /**
     * @param array $actEditorDocument
     * @return void
     */
    private function printAsString(array $actEditorDocument): void
    {
        $this->titleHorizontalAlignmentOnPrint = $actEditorDocument['titleHorizontalAlignmentOnPrint'];
        $this->section->addTextBreak(
            self::ONE_LINE_FOR_TEXT_BREAK,
            $this->predefinedStyles['sizeBreak'],
            $this->predefinedStyles['spaceBreak']
        );
        $this->table = $this->section->addTable();
        $this->table->addRow();
        switch ($this->titleHorizontalAlignmentOnPrint) {
            case 'enum':
                $this->printTableAsStringIfHorAlignmentEnum($actEditorDocument);
                break;
            case 'interlinearEnum':
                $this->printTableAsStringIfHorAlignmentInterLinearEnum($actEditorDocument);
                break;
            case 'interlinearInColumnEnum':
                $this->printTableAsStringIfHorAlignmentInterLinearInColumnEnum($actEditorDocument);
                break;
            case 'newlineInterlinearEnum':
                $this->printTableAsStringIfHorAlignmentNewLineInterLinearEnum($actEditorDocument);
        }
    }

    /**
     * @param array $actEditorDocument
     * @return void
     */
    private function printTableAsStringIfHorAlignmentEnum(array $actEditorDocument): void
    {
        [
            'actEditorDocument' => $actEditorDocument,
            'printStyle' => $printStyle,
            'tileText' => $tileText,
            'fieldText' => $fieldText,
            'isTitleBold' => $isTitleBold,
            'isTitleItalic' => $isTitleItalic,
        ] = $this->getDataForEnum($actEditorDocument);

        [
            'titleLengthInTwp' => $titleLengthInTwp,
            'fieldLengthInTwp' => $fieldLengthInTwp,
        ] = $this->getTitleAndFieldLengthsInTwp(
            $actEditorDocument,
            $tileText,
            $fieldText,
            $isTitleBold,
            $isTitleItalic
        );

        [
            'firstPartText' => $firstPartText,
            'secondPartText' => $secondPartText,
        ] = $this->getTextsParts($fieldLengthInTwp, $fieldText);

        if ($firstPartText === null) {
            $fieldLengthInTwp = $this->pageWidthInTwipWithoutMargins - $titleLengthInTwp;
        }

        $this->addCellForEnum(
            $titleLengthInTwp,
            $this->predefinedStyles["styleCellNoBorder"],
            $tileText,
            $printStyle
        );

        $this->addCellAfterTitleCell($firstPartText, $fieldText, $fieldLengthInTwp, $printStyle);

        $this->addCellOnSecondRowWithGridSpan(
            $firstPartText,
            $secondPartText,
            $titleLengthInTwp,
            $printStyle
        );
    }

    /**
     * @param array $actEditorDocument
     * @return array
     */
    private function getModifiedActEditorDocumentForEnum(array $actEditorDocument): array
    {
        $actEditorDocument['columnsAlign'] = $actEditorDocument['titleTypeSizeOnPrint'];
        $actEditorDocument['titleTypeSizeOnPrint'] = 'content';
        $actEditorDocument['fieldTypeSizeOnPrint'] = 'content';

        return $actEditorDocument;
    }

    /**
     * @param array $actEditorDocument
     * @return array
     */
    private function getDataForEnum(array $actEditorDocument): array
    {
        $actEditorDocument = $this->getModifiedActEditorDocumentForEnum($actEditorDocument);
        $printStyle = $this->getPrintStyle($actEditorDocument);
        $tileText = $actEditorDocument['fieldName'];
        $fieldText = is_null($this->model) ? '' : $this->printFieldValueColumn($actEditorDocument);
        $fontSize = $actEditorDocument['fontSize'] ?? $this->predefinedStyles['fontSize'];

        return [
            'actEditorDocument' => $actEditorDocument,
            'printStyle' => $printStyle,
            'tileText' => $tileText,
            'fieldText' => $fieldText,
            'fontSize' => $fontSize,
            'isTitleBold' => $printStyle['titleFontStyle']['bold'],
            'isTitleItalic' => $printStyle['titleFontStyle']['italic'],
        ];
    }

    /**
     * @param float $lengthInTwp
     * @param array $cellStyles
     * @param string|null $text
     * @param array $printStyle
     * @return void
     */
    private function addCellForEnum(
        float $lengthInTwp,
        array $cellStyles,
        ?string $text,
        array $printStyle
    ): void {
        $this->table
            ->addCell(
                $lengthInTwp,
                $cellStyles
            )
            ->addText(
                htmlspecialchars($text),
                $printStyle['titleFontStyle'],
                $printStyle['titleParagraphStyle']
            );
    }

    /**
     * @param string|null $firstPartText
     * @param string $fieldText
     * @param float $fieldLengthInTwp
     * @param array $printStyle
     * @return void
     */
    private function addCellAfterTitleCell(
        ?string $firstPartText,
        string $fieldText,
        float $fieldLengthInTwp,
        array $printStyle
    ): void {
        $text = $firstPartText ?? $fieldText;

        $this->addCellForEnum(
            $fieldLengthInTwp,
            $firstPartText
                ? $this->predefinedStyles["styleCellNoBorder"]
                : $this->predefinedStyles["styleCellLine"],
            $text,
            $printStyle
        );
    }

    /**
     * @param string|null $firstPartText
     * @param string|null $secondPartText
     * @param float $titleLengthInTwp
     * @param array $printStyle
     * @return void
     */
    private function addCellOnSecondRowWithGridSpan(
        ?string $firstPartText,
        ?string $secondPartText,
        float $titleLengthInTwp,
        array $printStyle
    ): void {
        $styles = $this->predefinedStyles["styleCellLine"];
        $styles['gridSpan'] = self::GRID_SPAN_COUNT;

        if ($firstPartText && $secondPartText) {
            $this->table->addRow();
            $this->addCellForEnum(
                $titleLengthInTwp,
                $styles,
                $secondPartText,
                $printStyle
            );
        }
    }

    /**
     * @param array $actEditorDocument
     * @param string|null $tileText
     * @param string|null $fieldText
     * @param bool $isTitleBold
     * @param bool $isTitleItalic
     * @return float[]|int[]|null[]
     */
    private function getTitleAndFieldLengthsInTwp(
        array $actEditorDocument,
        ?string $tileText,
        ?string $fieldText,
        bool $isTitleBold = false,
        bool $isTitleItalic = false
    ): array {
        $titleLengthInTwp = $this->calculateLengthOfTextInTwip(
            $actEditorDocument['titleTypeSizeOnPrint'],
            $actEditorDocument['titleProcentSizeOnPrint'] ?? null,
            $tileText,
            0,
            $isTitleBold,
            $isTitleItalic
        );
        $fieldLengthInTwp = $this->calculateLengthOfTextInTwip(
            $actEditorDocument['fieldTypeSizeOnPrint'],
            $actEditorDocument['fieldProcentSizeOnPrint'] ?? null,
            $fieldText,
        );

        return [
            'titleLengthInTwp' => $titleLengthInTwp,
            'fieldLengthInTwp' => $fieldLengthInTwp,
        ];
    }

    /**
     * @param float $fieldLengthInTwp
     * @param string $fieldText
     * @return array
     */
    private function getTextsParts(float $fieldLengthInTwp, string $fieldText): array
    {
        $firstPartText = null;
        $secondPartText = null;
        if ($fieldLengthInTwp > $this->pageWidthInTwipWithoutMargins) {
            $symbolLengthInTwp = $this->calculateSymbolLengthAndConvertToTwip();
            $twp = $fieldLengthInTwp - $this->pageWidthInTwipWithoutMargins;
            $number = $twp / $symbolLengthInTwp;
            $numberOfSymbols = mb_strlen($fieldText);
            $firstPartText = substr($fieldText, 0, $numberOfSymbols - $number);
            $secondPartText = substr($fieldText, $numberOfSymbols - $number);
        }

        return ['firstPartText' => $firstPartText, 'secondPartText' => $secondPartText];
    }

    /**
     * @param array $actEditorDocument
     * @return void
     */
    private function printTableAsStringIfHorAlignmentInterLinearEnum(
        array $actEditorDocument
    ): void {
        $this->printTableAsStringIfHorAlignmentEnum($actEditorDocument);
        [
            'actEditorDocument' => $actEditorDocument,
            'tileText' => $tileText,
            'fieldText' => $fieldText,
            'isTitleBold' => $isTitleBold,
            'isTitleItalic' => $isTitleItalic,
        ] = $this->getDataForEnum($actEditorDocument);

        [
            'titleLengthInTwp' => $titleLengthInTwp,
            'fieldLengthInTwp' => $fieldLengthInTwp,
        ] = $this->getTitleAndFieldLengthsInTwp(
            $actEditorDocument,
            $tileText,
            $fieldText,
            $isTitleBold,
            $isTitleItalic
        );

        $this->table->addRow();

        $styles = $this->predefinedStyles["styleCellNoBorder"];
        $styles['gridSpan'] = self::GRID_SPAN_COUNT;

        $this->addEmptySubCell($fieldLengthInTwp, $styles, $titleLengthInTwp);

        $this->addSubCellForEnum(
            $this->pageWidthInTwipWithoutMargins,
            $styles,
            $actEditorDocument['subscriptOnPrint']
        );
    }

    /**
     * @param float $lengthInTwp
     * @param array $styles
     * @param string|null $subString
     * @return void
     */
    private function addSubCellForEnum(float $lengthInTwp, array $styles, ?string $subString): void
    {
        $this->table
            ->addCell(
                $lengthInTwp,
                $styles
            )
            ->addText(
                htmlspecialchars($subString ?? ''),
                $this->subStyles,
                ['align' => 'center']
            );
    }

    /**
     * @param float $fieldLengthInTwp
     * @param array $styles
     * @param float $titleLengthInTwp
     * @return void
     */
    private function addEmptySubCell(float $fieldLengthInTwp, array &$styles, float $titleLengthInTwp): void
    {
        if ($fieldLengthInTwp < $this->pageWidthInTwipWithoutMargins) {
            unset($styles['gridSpan']);
            $this->addSubCellForEnum($titleLengthInTwp, $styles, '');
        }
    }

    /**
     * @param array $actEditorDocument
     * @return void
     */
    private function printTableAsStringIfHorAlignmentInterLinearInColumnEnum(
        array $actEditorDocument
    ): void {
        $actEditorDocument['columnsAlign'] = $actEditorDocument['titleTypeSizeOnPrint'];
        $actEditorDocument['titleTypeSizeOnPrint'] = 'content';
        $titleLines = $this->divideTextIntoLinesByPaperWidth(
            $actEditorDocument['fieldName'],
        );
        $titleLastLine = end($titleLines);
        $this->titleLastLineLenghtInTwip = $this->calculateLengthOfTextInTwip(
            $actEditorDocument['titleTypeSizeOnPrint'],
            $actEditorDocument['titleProcentSizeOnPrint'],
            $actEditorDocument['fontSize'],
            $titleLastLine,
        );
        $actEditorDocument['viewFieldOnPrint'] = true;
        $actEditorDocument['fieldTypeSizeOnPrint'] = 'procent';
        $actEditorDocument['fieldProcentSizeOnPrint'] = 100 * (($this->pageWidthInTwipWithoutMargins - $this->titleLastLineLenghtInTwip) / $this->pageWidthInTwipWithoutMargins);
        $this->printRecord($actEditorDocument);
    }

    /**
     * @param array $actEditorDocument
     * @return void
     */
    private function printTableAsStringIfHorAlignmentNewLineInterLinearEnum(
        array $actEditorDocument
    ): void {
        $actEditorDocument['columnsAlign'] = $actEditorDocument['titleTypeSizeOnPrint'];
        $actEditorDocument['titleTypeSizeOnPrint'] = 'endLine';
        $actEditorDocument['viewFieldOnPrint'] = true;
        $actEditorDocument['fieldTypeSizeOnPrint'] = 'endLine';
        $this->printRecord($actEditorDocument);
    }

    /**
     * @param ActEditorPrintWordTableDTO $actEditorPrintWordTableDTO
     * @return void
     */
    public function printTableWithActSignatures(ActEditorPrintWordTableDTO $actEditorPrintWordTableDTO): void
    {
        if (
            isset($actEditorPrintWordTableDTO->signatures_with_kcp, $actEditorPrintWordTableDTO->signatures)
            && ($actEditorPrintWordTableDTO->signatures_with_kcp->isNotEmpty(
                ) || $actEditorPrintWordTableDTO->signatures->isNotEmpty())
        ) {
            $this->massAssignmentOfPropertiesFromDTO($actEditorPrintWordTableDTO);
            $this->section->addTextBreak(
                self::FIVE_LINES_FOR_TEXT_BREAK,
                $this->predefinedStyles['sizeBreak'],
                $this->predefinedStyles['spaceBreak']
            );
            if ($this->signaturesWithKcp->isNotEmpty()) {
                Html::addHtml($this->section, $this->generateHTMLForTableWithSignsAndKCP());
            } elseif ($this->signatures->isNotEmpty()) {
                Html::addHtml($this->section, $this->generateHTMLForTableWithSigns(self::STYLE_CELL_WIDTH));
            }
        }
    }
}
