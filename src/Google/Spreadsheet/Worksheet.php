<?php
/**
 * Copyright 2013 Asim Liaquat
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
namespace Google\Spreadsheet;

use SimpleXMLElement;
use DateTime;

/**
 * Worksheet.
 *
 * @package    Google
 * @subpackage Spreadsheet
 * @author     Asim Liaquat <asimlqt22@gmail.com>
 */
class Worksheet
{
    /**
     * A worksheet xml object
     *
     * @var \SimpleXMLElement
     */
    private $xml;

    private $postUrl;

    private $editCellPostUrl;

    /**
     * Initializes the worksheet object.
     *
     * @param SimpleXMLElement $xml
     */
    public function __construct(SimpleXMLElement $xml)
    {
        $xml->registerXPathNamespace('gs', 'http://schemas.google.com/spreadsheets/2006');
        $this->xml = $xml;
    }

    /**
     * Get the worksheet id. Returns the full url.
     *
     * @return string
     */
    public function getId()
    {
        return $this->xml->id->__toString();
    }

    /**
     * Get the updated date
     *
     * @return DateTime
     */
    public function getUpdated()
    {
        return new DateTime($this->xml->updated->__toString());
    }

    /**
     * Get the title of the worksheet
     *
     * @return string
     */
    public function getTitle()
    {
        return $this->xml->title->__toString();
    }

    /**
     * Get the number of rows in the worksheet
     *
     * @return int
     */
    public function getRowCount()
    {
        $el = current($this->xml->xpath('//gs:rowCount'));
        return (int) $el->__toString();
    }

    /**
     * Get the number of columns in the worksheet
     *
     * @return int
     */
    public function getColCount()
    {
        $el = current($this->xml->xpath('//gs:colCount'));
        return (int) $el->__toString();
    }

    /**
     * Get the list feed of this worksheet
     *
     * @return \Google\Spreadsheet\ListFeed
     */
    public function getListFeed($reverse = false, $sort = 'column:timestamp', $max = null)
    {
        $res = ServiceRequestFactory::getInstance()->get($this->getListFeedUrl($reverse, $sort, $max));
        return new ListFeed($res);
    }

    /**
     * Get the cell feed of this worksheet
     *
     * @return \Google\Spreadsheet\CellFeed
     */
    public function getCellFeed($minRow = null, $maxRow = null, $minCol = null, $maxCol = null)
    {
        $res = ServiceRequestFactory::getInstance()->get($this->getCellFeedUrl($minRow, $maxRow, $minCol, $maxCol));
        return new CellFeed($res);
    }

    /**
     * Delete this worksheet
     *
     * @return null
     */
    public function delete()
    {
        ServiceRequestFactory::getInstance()->delete($this->getEditUrl());
    }

    public function setPostUrl($url)
    {
        $this->postUrl = $url;
    }

    /**
     * Get the edit url of the worksheet
     *
     * @return string
     */
    public function getEditUrl()
    {
        return Util::getLinkHref($this->xml, 'edit');
    }

    /**
     * The url which is used to fetch the data of a worksheet as a list
     *
     * @return string
     */
    public function getListFeedUrl($reverse = false, $sort = 'column:timestamp', $max = null)
    {
        $url = Util::getLinkHref($this->xml, 'http://schemas.google.com/spreadsheets/2006#listfeed');
        $query = array();
        if($reverse) {
          $query[] = "reverse=true";
        }

        if($sort) {
          $query[] = "sort=" . $sort;
        }

        if($max) {
          $query[] = "max-results=" . $max;
        }

        if(count($query)) {
          $url .= '?' . implode('&', $query);
        }
        return $url;
    }

    /**
     * Get the cell feed url
     *
     * @return stirng
     */
    public function getCellFeedUrl($minRow = null, $maxRow = null, $minCol = null, $maxCol = null)
    {
      $url = Util::getLinkHref($this->xml, 'http://schemas.google.com/spreadsheets/2006#cellsfeed');
      $params = array();
      if(isset($minRow)) {
        $params[] = "min-row=$minRow";
      }
      if(isset($maxRow)) {
        $params[] = "max-row=$maxRow";
      }
      if(isset($minCol)) {
        $params[] = "min-col=$minCol";
      }
      if(isset($maxCol)) {
        $params[] = "max-col=$maxCol";
      }
      if(!empty($params)) {
        $url .= '?' . implode("&", $params);
      }
      return $url;
    }
}

