<?php

namespace App\DTO\ActEditorPrint;

use App\DTO\DataTransferObject;
use App\Extensions\DTO\TransferToCamelCaseForDTOTrait;
use Illuminate\Support\Collection;
use PhpOffice\PhpWord\Element\Section;

/**
 * Class ActEditorPrintWordTableDTO
 * @package App\DTO\ActEditorPrint
 *
 * @author Kozy-Korpesh Tolep
 */
class ActEditorPrintWordRecordDTO extends DataTransferObject
{
    use TransferToCamelCaseForDTOTrait;

    /**
     * @var float
     */
    public float $page_width_in_twip_without_margins;

    /**
     * @var \PhpOffice\PhpWord\Element\Section
     */
    public Section $section;

    /**
     * @var \Illuminate\Support\Collection|null
     */
    public ?Collection $act_editor_documents;

    /**
     * @var array|null
     */
    public ?array $current_act_editor_document;

    /**
     * @var int|null
     */
    public ?int $key_of_current_act_editor_document;

    /**
     * @var object|null
     */
    public ?object $model;

    /**
     * @var object|null
     */
    public ?object $object;

}