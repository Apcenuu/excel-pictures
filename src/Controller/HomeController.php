<?php

namespace App\Controller;

use App\Service\ApiService;
use App\Service\ExcelService;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\Response;

class HomeController extends AbstractController
{
    public function index()
    {
        return $this->render('index.html.twig');
    }

    public function downloadOrders($count)
    {
        if ($count <= 20) {
            $count = 20;
        }
        if ($count > 20 && $count <= 100) {
            $count = 100;
        }
        if ($count > 100) {
            $count = 100;
        }

        $excelService = new ExcelService();
        $apiUrl = $this->getParameter('app.api_url');
        $apiKey = $this->getParameter('app.api_key');
        $apiService = new ApiService($apiUrl, $apiKey);
        $orders = $apiService->getOrders($count);
        $products = $apiService->getProductsByOrders($orders);

        $content = $excelService->generateXlsx($products, $orders);

        return new Response($content, 200, [
            'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            "Content-Disposition: attachment; filename=orders.xlsx"
        ]);
    }
}
