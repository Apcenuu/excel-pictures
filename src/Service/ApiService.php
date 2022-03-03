<?php

namespace App\Service;

use RetailCrm\Api\Factory\SimpleClientFactory;
use RetailCrm\Api\Model\Filter\Orders\OrderFilter;
use RetailCrm\Api\Model\Filter\Store\ProductFilterType;
use RetailCrm\Api\Model\Filter\Store\ProductGroupFilterType;
use RetailCrm\Api\Model\Request\Orders\OrdersRequest;
use RetailCrm\Api\Model\Request\Store\ProductGroupsRequest;
use RetailCrm\Api\Model\Request\Store\ProductsRequest;

class ApiService
{

    private $client;

    public function __construct($apiUrl, $apiKey)
    {
        $this->client = SimpleClientFactory::createClient($apiUrl, $apiKey);
    }

    public function getOrders($limit)
    {
        $request = new OrdersRequest();
        $filter = new OrderFilter();
        $filter->extendedStatus[] = 'client-confirmed';
        $filter->extendedStatus[] = 'assembling';
        $filter->extendedStatus[] = 'prepayed';
        $request->filter = $filter;
        $request->limit = $limit;
        $response = $this->client->orders->list($request);
        return $response->orders;
    }

    public function getProduct($id)
    {
        $request = new ProductsRequest();
        $filter = new ProductFilterType();
        $request->filter = $filter;
        $filter->offerIds = [$id];
        $response = $this->client->store->products($request);

        return array_shift($response->products);
    }

    public function getProducts()
    {
        $page = 1;
        $products = [];
        do {
            $request = new ProductsRequest();
            $filter = new ProductFilterType();
            $request->filter = $filter;
            $request->limit = 100;
            $request->page = $page;
            $response = $this->client->store->products($request);

            foreach ($response->products as $product) {
                $products[] = $product;
            }
            $page++;

        } while ($page <= $response->pagination->totalPageCount);

        return $products;
    }

    public function getProductGroupById($categoryId)
    {
        $request = new ProductGroupsRequest();
        $filter = new ProductGroupFilterType();
        $filter->ids = [$categoryId];
        $request->filter = $filter;
        $request->limit = 100;
        $response = $this->client->store->productGroups($request);
        $productGroup = array_shift($response->productGroup);

        $category = [
            'id' => $productGroup->id,
            'name' => htmlspecialchars($productGroup->name),
            'parentId' => $productGroup->parentId
        ];

        return $category;
    }

    public function getProductsByOrders(array $orders)
    {
        $products = [];
        foreach ($orders as $order) {
            foreach ($order->items as $item) {
                $product = $this->getProduct($item->offer->id);
                $products[$item->offer->id] = $product;
            }
        }
        return $products;
    }
}
