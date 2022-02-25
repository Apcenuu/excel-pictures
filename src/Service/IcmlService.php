<?php

namespace App\Service;

use App\Service\RetailcrmIcml;
use Psr\Container\ContainerInterface;

class IcmlService
{
    private ExcelService $excelService;
    private CategoryService $categoryService;
    private ContainerInterface $container;
    private ApiService $apiService;

    public function __construct(ExcelService $excelService, CategoryService $categoryService, ApiService $apiService = null)
    {
        $this->excelService = $excelService;
        $this->categoryService = $categoryService;
        if ($apiService) {
            $this->apiService = $apiService;
        }

    }

    public function generateIcmlByProductArray(array $products)
    {
        $offers = [];
        $categories = [];

        foreach ($products as $product) {
            $offer = (array) array_shift($product->offers);
            $categoryId = array_shift($product->groups)->id;

            $havingParentCategory = array_filter($categories, function ($category) use ($categoryId) {
                if ($category['id'] == $categoryId) {
                    return true;
                }
                return false;
            });
            if (count($havingParentCategory) == 0) {
                $categories[] = $this->apiService->getProductGroupById($categoryId);
            }

            $offer['productId'] = $product->id;
            $offer['categoryId'] = [$categoryId];
            $offer['price'] = 0;
            $offer['picture'] = $product->imageUrl;
            $offer['url'] = $product->url;

            $offers[] = $offer;
        }

        foreach ($categories as $key => $category) {
            $categoryParent = $category['parentId'];

            if ($categoryParent !== null) {
                $categories[] = $this->apiService->getProductGroupById($categoryParent);
            }
        }

        $resultFile = 'MerKomuna.xml';
        $xmlDir = 'xml/';
        $this->excelService->clearDirectory($xmlDir);

        $icml = new RetailcrmIcml('MerKomuna', $xmlDir . $resultFile);
        $icml->generate($categories, $offers);

    }

    public function generateIcmlByFile($fileName)
    {
        $rows = $this->excelService->readFile($fileName);

        $offers = [];
        $categoryIdCounter = 1;
        $categories = [
            [
                'id' => $categoryIdCounter,
                'name' => 'Main'
            ]
        ];
        $categoryIdCounter++;
//        dump($rows);die;
        foreach ($rows as $key => $row) {

            $havingRowCategory = array_filter($categories, function ($category) use ($row) {
                if ($category['name'] == $row[0]) {
                    return true;
                }
                return false;
            });

            if (count($havingRowCategory) == 0 && $row[0] !== null) {
                $categories[] = [
                    'id' => $categoryIdCounter,
                    'name' => $row[0]
                ];
                $categoryIdCounter++;
            }

//            foreach ($row as $cellKey => $cell) {
//                if (is_null($cell)) {
//                    unset($row[$cellKey]);
//                }
//            }

            $productName = ucfirst($this->getValueFromRow($row,1));
            $offer = [
                'id' => uniqid(),
//                'description' => $row[2],
                'categoryId' => [$categoryIdCounter-1],
                'productId' => uniqid(),
//                'parent' => htmlspecialchars(utf8_encode(array_shift($offerCategory))),
//                'category' => $offerCategory,
                'productName'=> $productName,
                'quantity' => 1,
                'name' => ucfirst($this->getValueFromRow($row,3)),
                'price' => $this->getPrice($this->getValueFromRow($row,4)),
//                'purchasePrice' => $this->getPrice($this->getValueFromRow($row,10)),
                'article' => $this->getValueFromRow($row,2),
                'url' => $this->getValueFromRow($row,7),
//                'color' => ucfirst($this->getValueFromRow($row,7)),
//                'size' => $this->getValueFromRow($row,8),
//                'vendor' => ucfirst($this->getValueFromRow($row, 3)),
//                'picture' => $this->getValueFromRow($row,6),
                'description' => $this->getValueFromRow($row, 3)
            ];
//            dump($row, $offer);die;
            $offers[] = $offer;
//
        }
//        dump($categories);die;
        $resultFile = 'MerKomuna.xml';
        $xmlDir = 'xml/';
        $this->excelService->clearDirectory($xmlDir);

        $icml = new RetailcrmIcml('MerKomuna', $xmlDir . $resultFile);
        $icml->generate($categories, $offers);

        return $resultFile;
    }

    private function getPrice($price)
    {
        if (is_null($price)) {
            return 0;
        }

        return (float) trim(str_replace('S/', '', $price));
    }

    private function resultFile(): string
    {
        return uniqid() . '.xml';
    }

    private function getValueFromRow($row, $index, bool $int = false)
    {
        if (!isset($row[$index])) {
            if ($int) {
                return 0;
            }
            return '';
        }
        if ($int) {
            return (int) $row[$index];
        }
        return (string) $row[$index];
    }
}
