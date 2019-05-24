<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPWord
 * @copyright   2010-2016 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Writer\Word2007\Part;

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpWord\Element\Object;
use PhpOffice\PhpWord\Element\Title;
use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\Element\Text;
use PhpOffice\PhpWord\Element\TOC;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\Element\Image;
use PhpOffice\PhpWord\Element\ListItem;
use PhpOffice\PhpWord\Writer\Word2007\Element\Container;
use PhpOffice\PhpWord\Element\Link;
use PhpOffice\PhpWord\Element\PageBreak;
use PhpOffice\PhpWord\Element\TextBreak;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\Writer\Word2007\Element\Title as TitleWriter;
use PhpOffice\PhpWord\Writer\Word2007\Element\Table as TableWriter;
use PhpOffice\PhpWord\Writer\Word2007\Element\Object as ObjectWriter;
use PhpOffice\PhpWord\Writer\Word2007\Element\Text as TextWriter;
use PhpOffice\PhpWord\Writer\Word2007\Element\TextRun as TextRunWriter;
use PhpOffice\PhpWord\Writer\Word2007\Element\Image as ImageWriter;
use PhpOffice\PhpWord\Writer\Word2007\Element\Link as LinkWriter;
use PhpOffice\PhpWord\Writer\Word2007\Element\ListItem as ListItemWriter;
use PhpOffice\PhpWord\Writer\Word2007\Element\ListItemRun as ListItemRunWriter;
use PhpOffice\PhpWord\Writer\Word2007\Element\TOC as TOCWriter;
use PhpOffice\PhpWord\Writer\Word2007\Element\TextBreak as TextBreakWriter;
use PhpOffice\PhpWord\Writer\Word2007\Element\PageBreak as PageBreakWriter;
use PhpOffice\PhpWord\Writer\Word2007\Style\Section as SectionStyleWriter;

/**
 * Word2007 document part writer: word/document.xml
 */
class Document extends AbstractPart
{
    /**
     * Write part
     *
     * @return string
     */
    public function write()
    {
        $phpWord = $this->getParentWriter()->getPhpWord();
        $xmlWriter = $this->getXmlWriter();

        $sections = $phpWord->getSections();
        $sectionCount = count($sections);
        $currentSection = 0;
        $drawingSchema = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

        $xmlWriter->startDocument('1.0', 'UTF-8', 'yes');
        $xmlWriter->startElement('w:document');
        $xmlWriter->writeAttribute('xmlns:ve', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
        $xmlWriter->writeAttribute('xmlns:o', 'urn:schemas-microsoft-com:office:office');
        $xmlWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $xmlWriter->writeAttribute('xmlns:m', 'http://schemas.openxmlformats.org/officeDocument/2006/math');
        $xmlWriter->writeAttribute('xmlns:v', 'urn:schemas-microsoft-com:vml');
        $xmlWriter->writeAttribute('xmlns:wp', $drawingSchema);
        $xmlWriter->writeAttribute('xmlns:w10', 'urn:schemas-microsoft-com:office:word');
        $xmlWriter->writeAttribute('xmlns:w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
        $xmlWriter->writeAttribute('xmlns:wne', 'http://schemas.microsoft.com/office/word/2006/wordml');

        $xmlWriter->startElement('w:body');


        if ($sectionCount > 0) {
            foreach ($sections as $section) {
                $currentSection++;

                $containerWriter = new Container($xmlWriter, $section);
                $containerWriter->write();

                if ($currentSection == $sectionCount) {
                    $this->writeSectionSettings($xmlWriter, $section);
                } else {
                    $this->writeSection($xmlWriter, $section);
                }
            }
        }

        $xmlWriter->endElement(); // w:body
        $xmlWriter->endElement(); // w:document

        return $xmlWriter->getData();
    }

    /**
     * Write begin section.
     *
     * @param \PhpOffice\Common\XMLWriter $xmlWriter
     * @param \PhpOffice\PhpWord\Element\Section $section
     * @return void
     */
    private function writeSection(XMLWriter $xmlWriter, Section $section)
    {
        $xmlWriter->startElement('w:p');
        $xmlWriter->startElement('w:pPr');
        $this->writeSectionSettings($xmlWriter, $section);
        $xmlWriter->endElement();
        $xmlWriter->endElement();
    }

    /**
     * Write end section.
     *
     * @param \PhpOffice\Common\XMLWriter $xmlWriter
     * @param \PhpOffice\PhpWord\Element\Section $section
     * @return void
     */
    private function writeSectionSettings(XMLWriter $xmlWriter, Section $section)
    {
        $xmlWriter->startElement('w:sectPr');

        // Header reference
        foreach ($section->getHeaders() as $header) {
            $rId = $header->getRelationId();
            $xmlWriter->startElement('w:headerReference');
            $xmlWriter->writeAttribute('w:type', $header->getType());
            $xmlWriter->writeAttribute('r:id', 'rId' . $rId);
            $xmlWriter->endElement();
        }

        // Footer reference
        foreach ($section->getFooters() as $footer) {
            $rId = $footer->getRelationId();
            $xmlWriter->startElement('w:footerReference');
            $xmlWriter->writeAttribute('w:type', $footer->getType());
            $xmlWriter->writeAttribute('r:id', 'rId' . $rId);
            $xmlWriter->endElement();
        }

        // Different first page
        if ($section->hasDifferentFirstPage()) {
            $xmlWriter->startElement('w:titlePg');
            $xmlWriter->endElement();
        }

        // Section settings
        $styleWriter = new SectionStyleWriter($xmlWriter, $section->getStyle());
        $styleWriter->write();

        $xmlWriter->endElement(); // w:sectPr
    }

    public function getObjectAsText($element)
    {
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }
        if ($element instanceof Section) {

        } elseif ($element instanceof Text) {
            $textwriter = new TextWriter($objWriter, $element);
            $textwriter->write();
            $objWriter = $textwriter->getXmlWriter();
        } elseif ($element instanceof TextRun) {
            $textrunwriter = new TextRunWriter($objWriter, $element);
            $textrunwriter->write();
            $objWriter = $textrunwriter->getXmlWriter();
        } elseif ($element instanceof Link) {
            $linkwriter = new LinkWriter($objWriter, $element);
            $linkwriter->write();
            $objWriter = $linkwriter->getXmlWriter();
        } elseif ($element instanceof Title) {
            $titlewriter = new TitleWriter($objWriter, $element);
            $titlewriter->write();
            $objWriter = $titlewriter->getXmlWriter();
        } elseif ($element instanceof TextBreak) {
            $textbreakwriter = new TextBreakWriter($objWriter, $element);
            $textbreakwriter->write();
            $objWriter = $textbreakwriter->getXmlWriter();
        } elseif ($element instanceof PageBreak) {
            $pagebreakwriter = new PageBreakWriter($objWriter, $element);
            $pagebreakwriter->write();
            $objWriter = $pagebreakwriter->getXmlWriter();
        } elseif ($element instanceof Table) {
            $tablewriter = new TableWriter($objWriter, $element);
            $tablewriter->write();
            $objWriter = $tablewriter->getXmlWriter();
        } elseif ($element instanceof ListItem) {
            $listitemwriter = new ListItemWriter($objWriter, $element);
            $listitemwriter->write();
            $objWriter = $listitemwriter->getXmlWriter();
        } elseif ($element instanceof Image) {
            $imagewriter = new ImageWriter($objWriter, $element);
            $imagewriter->write();
            $objWriter = $imagewriter->getXmlWriter();
        } elseif ($element instanceof Object) {
            $objectwriter = new ObjectWriter($objWriter, $element);
            $objectwriter->write();
            $objWriter = $objectwriter->getXmlWriter();
        } elseif ($element instanceof TOC) {
            $tocwriter = new TOCWriter($objWriter, $element);
            $tocwriter->write();
            $objWriter = $tocwriter->getXmlWriter();
        }
        return trim($objWriter->getData());
    }
}
