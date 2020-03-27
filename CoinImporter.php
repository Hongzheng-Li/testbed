<?php

namespace Topxia\Service\Importer;

use Topxia\Common\FileToolkit;
use Topxia\Common\ArrayToolkit;
use Symfony\Component\HttpFoundation\Request;

class CoinImporter extends Importer
{
    protected $necessaryFields = array('id'=>'ID' , 'verifiedMobile' => '手机号' , 'amount' => '虚拟币金额' , 'name' => '原因' , 'category' => '类别');
    protected $objWorksheet;
    protected $rowTotal         = 0;
    protected $colTotal         = 0;
    protected $excelFields      = array();
    protected $passValidateUser = array();

    protected $type = 'coin';

    public function import(Request $request)
    {
        $importData = $request->request->get('importData');

        return $this->excelDataImporting([], $importData);
    }


    protected function excelDataImporting($targetObject, $coinData)
    {
        $existsUserCount = 0;
        $successCount    = 0;

        $userCount = 0;
        $amountCount = 0;
        //开始添加数据
        foreach ($coinData as $key => $coin) {
            foreach($coin as $k=>$v){
                $coin[$k] = trim($v);
            }

            if (!empty($coin['id'])) {
                $user = $this->getUserService()->getUser($coin['id']);
            } elseif (!empty($coin['verifiedMobile'])) {
                $user = $this->getUserService()->getUserByVerifiedMobile($coin['verifiedMobile']);
            }

            $flow = [
                'userId'=>$user['id'],
                'amount'=>$coin['amount'],
                'name'=>$coin['name'],
                'orderSn'=>'',
                'category'=>$coin['category'],
                'note'=>''
            ];
            $this->getCashService()->inflowByCoin($flow);
            $userCount += 1;
            $amountCount += $coin['amount'];
        }
        $this->getLogService()->info('coin','import',sprintf('一共给%s个学员，共添加了%s虚拟币',$userCount,$amountCount),$coinData);
        return array('existsUserCount' => $existsUserCount, 'successCount' => $successCount);
    }

    public function check(Request $request)
    {
        $file     = $request->files->get('excel');

        $danger   = $this->validateExcelFile($file);
        if (!empty($danger)) {
            return $danger;
        }

        $importData = $this->getCoinData();

        if (!empty($importData['errorInfo'])) {
            return $this->createErrorResponse($importData['errorInfo']);
        }

        return $this->createSuccessResponse(
            $importData['allUserData'],
            $importData['checkInfo'],
            array(

            ));
    }

    protected function getCoinData()
    {
        $fieldSort   = $this->getFieldSort();
        $validate    = array();

        for ($row = 3; $row <= $this->rowTotal; $row++) {
            for ($col = 0; $col < $this->colTotal; $col++) {
                $infoData          = $this->objWorksheet->getCellByColumnAndRow($col, $row)->getFormattedValue();
                $columnsData[$col] = $infoData."";
            }
            foreach ($fieldSort as $sort) {
                $coinData[$sort['fieldName']] = $columnsData[$sort['num']];
                $fieldCol[$sort['fieldName']] = $sort['num'] + 1;
            }

            $emptyData = array_count_values($coinData);

            if (isset($emptyData[""]) && count($coinData) == $emptyData[""]) {
                $checkInfo[] = $this->getServiceKernel()->trans('第%row%行为空行，已跳过', array('%row%' => $row));
                continue;
            }

            $info                            = $this->validExcelFieldValue($coinData, $row, $fieldCol);
            empty($info) ? '' : $errorInfo[] = $info;

            if (empty($errorInfo)) {
                $validate[] = array_merge($coinData, array('row' => $row));
            }
            unset($coinData);
        }

        $this->passValidateUser = $validate;

        $data['errorInfo']   = empty($errorInfo) ? array() : $errorInfo;
        $data['checkInfo']   = empty($checkInfo) ? array() : $checkInfo;
        $data['allUserData'] = empty($this->passValidateUser) ? array() : $this->passValidateUser;

        return $data;
    }

    protected function validateExcelFile($file)
    {
        if (!is_object($file)) {
            return $this->createDangerResponse($this->getServiceKernel()->trans('请选择上传的文件'));
        }

        if (FileToolkit::validateFileExtension($file, 'xls xlsx')) {
            return $this->createDangerResponse($this->getServiceKernel()->trans('Excel格式不正确！'));
        }

        //获取excel信息
        $this->excelAnalyse($file);

        if ($this->rowTotal > 1000) {
            return $this->createDangerResponse($this->getServiceKernel()->trans('Excel超过1000行数据!'));
        }

        if (!$this->checkNecessaryFields($this->excelFields)) {
            return $this->createDangerResponse($this->getServiceKernel()->trans('缺少必要的字段'));
        }
    }

    protected function validExcelFieldValue($coinData, $row, $fieldCol)
    {
        foreach($coinData as $k=>$v){
            $coinData[$k] = trim($v);
        }
        $errorInfo = '';
        $user = [];

        if (!empty($coinData['id'])) {
            $user = $this->getUserService()->getUser($coinData['id']);
        } elseif (!empty($coinData['verifiedMobile'])) {
            $user = $this->getUserService()->getUserByVerifiedMobile($coinData['verifiedMobile']);
        }


        if (empty($user) || (empty($coinData['id']) && empty($coinData['verifiedMobile']))) {
            $user = null;
        }

        //没有用户
        if (!$user) {
            $errorInfo = $this->getServiceKernel()->trans('第 %row%行的信息有误，用户数据不存在，请检查。', array('%row%' => $row));
        }else{
            foreach($this->necessaryFields as $k=>$value){
                //数据没有填写完整
                if($k != 'id' && $k != 'verifiedMobile'){
                    if(empty($coinData[$k])){
                        $errorInfo = $this->getServiceKernel()->trans('第 %row%行的信息有误，仔细看提示！填写完成信息。', array('%row%' => $row));
                    }else{
                        if($k == 'amount'){
                            if(!is_numeric($coinData[$k])){
                                $errorInfo = $this->getServiceKernel()->trans('第 %row%行的信息有误，虚拟币金额必须为整数类型！。', array('%row%' => $row));
                            }else{
                                if($coinData[$k]<=0){
                                    $errorInfo = $this->getServiceKernel()->trans('第 %row%行的信息有误，虚拟币金额必须大于0。', array('%row%' => $row));
                                }
                                if(!preg_match("/^[1-9][0-9]*$/" ,$coinData[$k])){
                                    $errorInfo = $this->getServiceKernel()->trans('第 %row%行的信息有误，虚拟币金额不能为小数。', array('%row%' => $row));
                                }
                            }
                        }
                    }
                }
            }
        }
        return $errorInfo;
    }



    protected function arrayRepeat($array, $nickNameCol)
    {
        $repeatArrayCount = array_count_values($array);
        $repeatRow        = "";

        foreach ($repeatArrayCount as $key => $value) {
            if ($value > 1 && !empty($key)) {
                $repeatRow .= $this->getServiceKernel()->trans('第%col%列重复:', array('%col%' => ($nickNameCol + 1))).'<br>';

                for ($i = 1; $i <= $value; $i++) {
                    $row = array_search($key, $array) + 3;

                    $repeatRow .= $this->getServiceKernel()->trans('第%row%行    %key%', array('%row%' => $row, '%key%' => $key)).'<br>';

                    unset($array[$row - 3]);
                }
            }
        }

        return $repeatRow;
    }

    protected function getFieldSort()
    {
        $fieldSort       = array();
        $necessaryFields = $this->getNecessaryFields();
        $excelFields     = $this->excelFields;

        foreach ($excelFields as $key => $value) {
            if (in_array($value, $necessaryFields)) {
                foreach ($necessaryFields as $fieldKey => $fieldValue) {
                    if ($value == $fieldValue) {
                        $fieldSort[$fieldKey] = array("num" => $key, "fieldName" => $fieldKey);
                        break;
                    }
                }
            }
        }
        return $fieldSort;
    }

    protected function excelAnalyse($file)
    {
        $objPHPExcel        = \PHPExcel_IOFactory::load($file);
        $objWorksheet       = $objPHPExcel->getActiveSheet();
        $highestRow         = $objWorksheet->getHighestRow();
        $highestColumn      = $objWorksheet->getHighestColumn();
        $highestColumnIndex = \PHPExcel_Cell::columnIndexFromString($highestColumn);
        $excelFields        = array();

        for ($col = 0; $col < $highestColumnIndex; $col++) {
            $fieldTitle                                  = $objWorksheet->getCellByColumnAndRow($col, 2)->getValue();
            empty($fieldTitle) ? '' : $excelFields[$col] = $this->trim($fieldTitle);
        }

        $rowAndCol = array('rowLength' => $highestRow, 'colLength' => $highestColumnIndex);

        $this->objWorksheet = $objWorksheet;
        $this->rowTotal     = $highestRow;
        $this->colTotal     = $highestColumnIndex;
        $this->excelFields  = $excelFields;
        return array($objWorksheet, $rowAndCol, $excelFields);
    }

    protected function checkNecessaryFields($excelFields)
    {
        return ArrayToolkit::some($this->necessaryFields, function ($fields) use ($excelFields) {
            return in_array($fields, array_values($excelFields));
        });
    }

    public function getTemplate(Request $request)
    {
        return $this->render('TopxiaAdminBundle:Coin:import.html.twig', array(
            'importerType' => $this->type
        ));
    }

    public function tryImport(Request $request)
    {

    }

    public function getNecessaryFields()
    {
        $necessaryFields = array('id'=>'ID' , 'verifiedMobile' => '手机号' , 'amount' => '虚拟币金额' , 'name' => '原因' , 'category' => '类别');
        return $this->getServiceKernel()->transArray($necessaryFields);
    }

    protected function getUserService()
    {
        return $this->getServiceKernel()->createService('User.UserService');
    }

    protected function getLogService()
    {
        return $this->getServiceKernel()->createService('System.LogService');
    }

    protected function getCashService()
    {
        return $this->getServiceKernel()->createService('Cash.CashService');
    }
}
