<?php

namespace App\Service;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx\Worksheet;
use RetailCrm\Api\Model\Entity\Orders\Order;
use RetailCrm\Api\Model\Entity\Store\Product;

class ExcelService
{
    const MAX_FILES = 1;

    public function readFile($filename)
    {
        $reader = new Xlsx();
        $spreadsheet = $reader->load($filename);
        $worksheet = $spreadsheet->getActiveSheet();
        $rows = $worksheet->toArray();
        unset($rows[0]);
        return $rows;
    }

    public function clearDirectory($directory)
    {
        $files = array_filter(scandir($directory), function ($value, $key) {
            $fileExtensionArray = explode('.', $value);
            $fileExtension = end($fileExtensionArray);
            if ($fileExtension === 'xlsx' || $fileExtension === 'xml') {
                return true;
            }
            return false;
        }, ARRAY_FILTER_USE_BOTH);
        $countFiles = count($files);

        while ($countFiles >= self::MAX_FILES) {
            unlink($directory . '/'. array_shift($files));
            $countFiles = count($files);
        }
    }

    public function generateXlsx($products, $orders)
    {
        $filename = "template.xlsx";
        $rows[] = [
            'A' => 'Cliente',
            'B' => 'Artículo',
            'C' => 'Tipo de entrega',
            'D' => 'Ciudad',
            'E' => 'Dirección de entrega',
            'F' => 'Coste de entrega',
            'G' => 'Tipo de pago',
            'H' => 'Importe/Cantidad', //Price
            'I' => 'Pagado', //Amount
            'J' => 'Ficha tecnica del articulo',
            'L' => 'Tienda',
            'M' => 'Fecha de entrega', //Delivery date
            'N' => 'Descuento (%)', //Discount
            'O' => 'Fecha de pago', //Payment Date
            'P' => 'Cantidad a pagar', //Остаток оплаты
            'Q' => 'Adjuntos', //Picture
        ];
        $spreadsheet = new Spreadsheet();
        $number = 0;
        /** @var Product $product */
        foreach ($products as $orderId => $product) {
            $order = array_filter($orders, function ($order) use ($orderId) {
                if ($order->id == $orderId) {
                    return true;
                }
                return false;
            });
            /** @var Order $order */
            $order = array_shift($order);
            $payment = array_shift($order->payments);
            $offer = array_shift($product->offers);
            $deliveryDate = null;
            if ($order->delivery->date) {
                $deliveryDate = $order->delivery->date->format('d.m.y H:i:s');
            }
            $paymentDate = null;
            if ($payment->paidAt) {
                $paymentDate = $payment->paidAt->format('d.m.y H:i:s');
            }


//            dump($payment);die;
            $row = [
                'A' => $order->customer->firstName . ' ' . $order->customer->lastName,
                'B' => $product->article,
                'C' => $order->delivery->code,
                'D' => $order->delivery->address->city,
                'E' => $order->delivery->address->text,
                'F' => $order->delivery->cost,
                'G' => $payment->type ?? null,
                'H' => $offer->price,
                'I' => $payment->amount ?? null,
                'J' => $order->customFields['fichatecnica'] ?? null,
                'L' => $order->customer->site,
                'M' => $deliveryDate, //Delivery date
                'N' => $order->discountManualPercent, //Discount
                'O' => $paymentDate, //Payment Date
                'P' => $order->purchaseSumm //Остаток оплаты

            ];
            $rowCount = 2 + $number;
            $spreadsheet = $this->addImageToXlsx($spreadsheet, $product->imageUrl, 'Q' . $rowCount);

            $rows[] = $row;
            $number++;
        }


        $spreadsheet->getActiveSheet()->fromArray($rows);
        foreach ($row as $key => $value) {
            $spreadsheet->getActiveSheet()->getColumnDimension($key)->setAutoSize(true);
        }

        $writer = IOFactory::createWriter($spreadsheet, "Xlsx");
        $writer->save($filename);

        return file_get_contents($filename);
    }

    private function addImageToXlsx(Spreadsheet $spreadsheet, string $filename, string $coordinate)
    {
        $fileExtension = explode('.', $filename);
        $fileExtension = end($fileExtension);
        $fileExtension = explode('?', $fileExtension);
        $fileExtension = array_shift($fileExtension);

        switch ($fileExtension) {
            case 'jpg':
                $gdImage = imagecreatefromjpeg($filename);
                break;
            case 'png':
                $gdImage = imagecreatefrompng($filename);
                break;
        }

        $row = substr($coordinate, 1);
        $column = substr($coordinate, 0, 1);

        $drawingObject = new MemoryDrawing();
        $drawingObject->setName('Sample image');
        $drawingObject->setDescription('Sample image');
        $drawingObject->setImageResource($gdImage);
        $drawingObject->setRenderingFunction(MemoryDrawing::RENDERING_JPEG);
        $drawingObject->setMimeType(MemoryDrawing::MIMETYPE_DEFAULT);
        $drawingObject->setHeight(150);
        $drawingObject->setWorksheet($spreadsheet->getActiveSheet());
        $drawingObject->setCoordinates($coordinate);

        $spreadsheet->getActiveSheet()->getRowDimension($row)->setRowHeight(150);
        $spreadsheet->getActiveSheet()->getColumnDimension($column)->setWidth(20);

        return $spreadsheet;
    }
}
