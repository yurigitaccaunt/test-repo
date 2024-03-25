<?php
namespace Fusion\Components;

use Bitrix\Main\Engine\Contract\Controllerable;
use Bitrix\Main\UI;
use Bitrix\Main\Grid;
use Fusion\Timesheet\ReportTable as TimeSheet;
use Fusion\Timesheet\Dictionaries;
use Fusion\Timesheet\Permission;
use Fusion\Core\System;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Writer\Common\Creator\Style\StyleBuilder;
use OpenSpout\Writer\Common\Creator\WriterEntityFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Fusion\Timeman;
use Fusion\Project\EntityTable as ProjectEntity;

class TimeSheetRegistryCommentComponent extends \CBitrixComponent implements Controllerable
{



    const GRID_ID = 'TIMESHEET_REGISTRY_V2';
    
    /**
     * Количество элементов на странице
     *
     * @var integer
     */
    const PAGE_SIZE = 20;
    
    public function configureActions()
    {
        return [];
    }
    
    public function executeComponent()
    {
        UI\Extension::load('jquery');
        
        $this->arResult['permissions'] = [
            'can_generate_main_report' => Permission::canGenerateMainReport(),
            'can_generate_detail_report' => Permission::canGenerateDetailReport(),
            'can_create_report' => Permission::isEmploye()
        ];
        
        $this->arResult['GRID_ID'] = self::GRID_ID . '_' . System::getUser()->getID();
        
        $this->arResult['GRID_FILTER'] = $this->getFilterFields();
        $this->arResult['COLUMNS'] = $this->getColumns();
        $this->arResult['PRESETS'] = $this->getFilterDefaultPesets();
        
        $entityObject = new TimeSheet();
        
        $grid = new Grid\Options($this->arResult['GRID_ID']);
        
        $navParams = $grid->GetNavParams();
        $nav = new UI\PageNavigation($this->arResult['GRID_ID']);
        
        $nav->allowAllRecords(true)
        ->setPageSize($navParams['nPageSize'] ?? self::PAGE_SIZE)
        ->initFromUri();
        
        $filter = $this->prepairFilter($this->arResult['GRID_ID'], $this->arResult['GRID_FILTER']);
        
        $filter = $this->getPermissionFilter($filter);
        
        $sort = $grid->GetSorting([
            'sort' => [
                'ID' => 'DESC'
            ],
            'vars' => [
                'by' => 'by',
                'order' => 'order'
            ]
        ]);
        
        $elementCollection = $entityObject::getList([
            'filter' => $filter,
            'select' => [
                '*',
                'DEPARTMENT_NAME' => 'DEPARTMENT.NAME'
            ],
            "order" => $sort['sort'],
            "count_total" => true,
            "offset" => $nav->getOffset(),
            "limit" => $nav->getLimit()
        ]);
        
        $nav->setRecordCount($elementCollection->getCount());
        
        foreach ($elementCollection as $element) {
            
            $data = $element;
            
            $element['RESPONSIBLE_ID'] = $this->getUserViewHtml($element['RESPONSIBLE_ID']);
            $element['STATUS'] = Dictionaries::getStatuses()[$element['STATUS']];
            
            $element['STATS_TOTAL'] = $this->getTimeStringBySeconds($element['STATS_TOTAL']);
            $element['STATS_TASK'] = $this->getTimeStringBySeconds($element['STATS_TASK']);
            $element['STATS_EVENT'] = $this->getTimeStringBySeconds($element['STATS_EVENT']);
            $element['STATS_NORM'] = $this->getTimeStringBySeconds($element['STATS_NORM']);
            
            $actions = [];
            
            $actions[] = [
                "TEXT" => "Открыть",
                "ONCLICK" => "showReport({$element['ID']})",
                'DEFAULT' => true
                ];
            
            $row = [
                'id' => $element['ID'],
                'data' => $data,
                'columns' => $element,
                'editable' => 'Y',
                'actions' => $actions
            ];
            
            $this->arResult['ROWS'][] = $row;
        }
        
        $this->arResult['NAV'] = $nav;
        
        $this->includeComponentTemplate();
    }
    
    public function generateMainReportAction()
    {
        if (! Permission::canGenerateMainReport()) {
            throw new \Exception('Недостаточно прав для генерации отчёта');
        }
        
        set_time_limit(1800);
        
        $header = [
            'ФИО сотрудника',
            'Подразделение',
            'Дата ТШ',
            'Статус ТШ',
            'Итого',
            'Время по задачам',
            'Время по событиям'
        ];
        
        ob_start();
        
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        
        $sheet->fromArray([
            $header
        ], NULL, 'A1');
        
        $styleArrayFirstRow = [
            'font' => [
                'bold' => true,
            ]
        ];
        
        $highestColumn = $sheet->getHighestColumn();
        $sheet->getStyle('A1:' . $highestColumn . '1' )->applyFromArray($styleArrayFirstRow);
        
        $entityObject = new TimeSheet();
        
        $grid_id = self::GRID_ID . '_' . System::getUser()->getID();
        $grid = new Grid\Options($grid_id);
        
        $filter = $this->prepairFilter($grid_id, $this->getFilterFields());
        
        $allowed = [
            '>=DATE_REPORT',
            '<=DATE_REPORT'
        ];
        
        $filter = array_filter($filter, function ($key) use ($allowed) {
            return in_array($key, $allowed);
        }, ARRAY_FILTER_USE_KEY);
            
            $sort = $grid->GetSorting([
                'sort' => [
                    'DATE_REPORT' => 'DESC'
                ],
                'vars' => [
                    'by' => 'by',
                    'order' => 'order'
                ]
            ]);
            
            $elements = [];
            
            $elementCollection = $entityObject::getList([
                'filter' => $filter,
                'select' => [
                    'STATS_EVENT',
                    'STATS_TASK',
                    'STATS_TOTAL',
                    'STATUS',
                    'DATE_REPORT',
                    'DEPARTMENT_NAME' => 'DEPARTMENT.NAME',
                    'RESPONSIBLE_NAME' => 'RESPONSIBLE.NAME',
                    'RESPONSIBLE_SECOND_NAME' => 'RESPONSIBLE.SECOND_NAME',
                    'RESPONSIBLE_LAST_NAME' => 'RESPONSIBLE.LAST_NAME'
                ],
                "order" => $sort['sort']
            ]);
            
            foreach ($elementCollection as $element) {
                
                $elements[] = [
                    implode(' ', array_filter([
                        $element['RESPONSIBLE_LAST_NAME'],
                        $element['RESPONSIBLE_NAME'],
                       // $element['RESPONSIBLE_SECOND_NAME']
                    ])),
                    $element['DEPARTMENT_NAME'],
                    $element['DATE_REPORT'],
                    Dictionaries::getStatuses()[$element['STATUS']],
                    $this->getTimeStringBySeconds($element['STATS_TOTAL']),
                    $this->getTimeStringBySeconds($element['STATS_TASK']),
                    $this->getTimeStringBySeconds($element['STATS_EVENT'])
                ];
            }
            
            $sheet->fromArray($elements, NULL, 'A2');
            
            // autosize
            foreach (range('A', $sheet->getHighestDataColumn()) as $col) {
                $sheet->getColumnDimension($col)->setAutoSize(true);
            }

            $fileName = 'main_report' . time() . '.xlsx';
            $filePath = sys_get_temp_dir() . '/' . $fileName;

            $writer = new Xlsx($spreadsheet);
            $writer->save($filePath);

            $arFile = \CFile::MakeFileArray($filePath);
            $fileId = \CFile::SaveFile($arFile, 'timesheet');

            return \CFile::GetPath($fileId);
    }
    
    public $monthLabels = [
        1  => 'Январь',
        2  => 'Февраль',
        3  => 'Март',
        4  => 'Апрель',
        5  => 'Май',
        6  => 'Июнь',
        7  => 'Июль',
        8  => 'Август',
        9  => 'Сентябрь',
        10  => 'Октябрь',
        11  => 'Ноябрь',
        12  => 'Декабрь',
    ];
    
    public function generateDetailReportAction()
    {
        if (! Permission::canGenerateDetailReport()) {
            throw new \Exception('Недостаточно прав для генерации отчёта');
        }
        
        set_time_limit(1800);
        
        $header = [
            'department' => 'Подразделение',
            'responsible' => 'ФИО сотрудника',
            'month_label' => 'Месяц',
            'date_report' => 'Дата ТШ',
            'direction' => 'Направление',
            'project_name' => 'Проект',
            'group_name' => 'Группа задач',
            'entity_type' => 'Вид деятельности',
            'title' => 'Деятельность',
            'date_from' => 'Дата начала',
            'date_to' => 'Дата завершения',
            'estimate' => 'Выделено',
            'elapsed' => 'Затрачено',
            'project_status' => 'Статус',
            'project_client' => 'Клиент',
            'project_group_number' => 'ГП',
            'project_curator_name' => 'Куратор проекта',
            'project_supervisor_name' => 'Руководитель проекта',
            'project_date_plan_start' => 'Дата старта проекта',
            'project_date_plan_finish' => 'Дата завершения проекта',
            'nrv_r' => 'НРВ Сотрудник',
            'nrv_d' => 'НРВ Подразделение',
            'workload_r' => '% Загрузки Сотрудник.',
            'workload_d' => '% Загрузки Подразделение.',
            'abs_nrv_r' => 'НРВ ("-" бол./отп.) Сотрудник.',
            'abs_nrv_d' => 'НРВ ("-" бол./отп.) Подразделение.',
            'abs_workload_r' => '% Загрузки (без учета бол/отп) Сотрудник.',
            'abs_workload_d' => '% Загрузки (без учета бол/отп) Подразделение.'
        ];
        
        $entityObject = new TimeSheet();
        
        $grid_id = self::GRID_ID . '_' . System::getUser()->getID();
        $grid = new Grid\Options($grid_id);
        
        $filter = $this->prepairFilter($grid_id, $this->getFilterFields());
        
        $allowed = [
            '>=DATE_REPORT',
            '<=DATE_REPORT'
        ];
        
        $filter = array_filter($filter, function ($key) use ($allowed) {
            return in_array($key, $allowed);
        }, ARRAY_FILTER_USE_KEY);
            
            $date_report_from_query = "";
            $date_report_to_query = "";
            
            if (! empty($filter['>=DATE_REPORT'])) {
                $date_object = new \Bitrix\Main\Type\DateTime($filter['>=DATE_REPORT']);
                $date_report_from_query = "AND report.DATE_REPORT >= '" . $date_object->format('Y-m-d H:i:s') . "'";
            }
            
            if (! empty($filter['<=DATE_REPORT'])) {
                $date_object = new \Bitrix\Main\Type\DateTime($filter['<=DATE_REPORT']);
                $date_report_to_query = "AND report.DATE_REPORT <= '" . $date_object->format('Y-m-d H:i:s') . "'";
            }
            
            // @todo Оптимизировать при больших нагрузках
            $sql = "SELECT 
            result.*, 
            user.`NAME`, user.`LAST_NAME`, user.`SECOND_NAME`,
            section.`NAME` as DEPARTMENT_NAME, 
            projects.ID as `PROJECT_ID`,
            projects.TITLE as `PROJECT_NAME`
            FROM (
            SELECT tst.`ID`, report.`DATE_REPORT`, report.`RESPONSIBLE_ID`, report.`DEPARTMENT_ID`, report.`STATS_NORM`,
            task.`TITLE` as `TITLE`, task.`START_DATE_PLAN` as `DATE_FROM`, task.`END_DATE_PLAN` as `DATE_TO`,
            tst.`REPORT_ID`, tst.`DIRECTION`, tst.`GROUP`, tst.`PROJECT_ID`, 'TASK' as `TYPE`, coalesce(elapsed.`SECONDS`, 0) as `ELAPSED`, `TIME_ESTIMATE`
            FROM f_ts_task tst
            LEFT JOIN b_tasks as task
            ON task.ID = tst.TASK_ID
            LEFT JOIN b_tasks_elapsed_time as elapsed
            ON elapsed.ID = tst.ELAPSED_ID
            LEFT JOIN f_ts_report report
            ON report.ID = tst.REPORT_ID
            WHERE report.STATUS != 'DRAFT' {$date_report_from_query} {$date_report_to_query}
            UNION
            SELECT tse.`ID`, report.`DATE_REPORT`, report.`RESPONSIBLE_ID`, report.`DEPARTMENT_ID`, report.`STATS_NORM`, event.`NAME` as `TITLE`, event.`DATE_FROM` as `DATE_FROM`, event.`DATE_TO` as `DATE_TO`,
            tse.`REPORT_ID`, tse.`DIRECTION`, tse.`GROUP`, tse.`PROJECT_ID`, 'EVENT' as `TYPE`, tse.`ELAPSED_VALUE` as `ELAPSED`, 0 as `TIME_ESTIMATE`
            FROM f_ts_event tse
            LEFT JOIN b_calendar_event as event
            ON event.ID = tse.EVENT_ID
            LEFT JOIN f_ts_report report
            ON report.ID = tse.REPORT_ID
            WHERE report.STATUS != 'DRAFT' {$date_report_from_query} {$date_report_to_query}
            UNION
            SELECT tsa.`ID`, report.`DATE_REPORT`, report.`RESPONSIBLE_ID`, report.`DEPARTMENT_ID`, report.`STATS_NORM`, tsa.`TYPE` as `TITLE`, tsa.`DATE_ABSENCE` as `DATE_FROM`, tsa.`DATE_ABSENCE` as `DATE_TO`,
            tsa.`REPORT_ID`, tsa.`DIRECTION`, tsa.`GROUP`, tsa.`PROJECT_ID`, 'EVENT_ABS' as `TYPE`, tsa.`ELAPSED_VALUE` as `ELAPSED`, 0 as `TIME_ESTIMATE`
            FROM f_ts_absence tsa
            LEFT JOIN f_ts_report report
            ON report.ID = tsa.REPORT_ID
            WHERE report.STATUS != 'DRAFT' {$date_report_from_query} {$date_report_to_query}
            ) as result
            LEFT JOIN b_user user
            ON result.RESPONSIBLE_ID = user.ID
            LEFT JOIN b_iblock_section section
            ON result.DEPARTMENT_ID = section.ID
            LEFT JOIN f_projects projects
            ON result.PROJECT_ID = projects.ID
            ORDER BY DATE_REPORT DESC, REPORT_ID DESC";

            $connection = System::getConnection();
            $query = $connection->query($sql);
            
            $directions = Dictionaries::getDirections();
            $dublicatebleDirections = Dictionaries::getDublicatebleDirections();
            $projectDirections = Dictionaries::getProjectDirections();
            
            $groups = Dictionaries::getGroups();
            $progectStatuses = ProjectEntity::getStatusList();

            $groupForAbscenseId = Dictionaries::getGroupForAbscenseId();
          
            $i = 0;
            
            //id проектов, которые встретились в данных
            $projectIds = [];

            foreach ($query as $row) {

                if($row['TYPE'] == 'EVENT_ABS')
                {
                    $row['TITLE'] = Dictionaries::getAbsenceTypes()[$row['TITLE']];
                }
                
                $responsibleName = implode(' ', array_filter([
                    $row['LAST_NAME'],
                    $row['NAME'],
                    //$row['SECOND_NAME']
                ]));
                
                $monthNumber = $row['DATE_REPORT']->format('n');
                $dateIndex = $row['DATE_REPORT']->format('Y.m');
                
                $projectName = '';
                
                if(in_array($row['DIRECTION'], $projectDirections))
                {
                    $projectName = $row['PROJECT_NAME'];
                }
                elseif(in_array($row['DIRECTION'], $dublicatebleDirections))
                {
                    $projectName = $directions[$row['DIRECTION']];
                }

                $element = [
                    'department'=>$row['DEPARTMENT_NAME'],
                    'responsible' => $responsibleName,
                    'month_label'=>$this->monthLabels[$monthNumber],
                    'date_report' => $row['DATE_REPORT']->format('d-m-Y'),
                    'direction' => $directions[$row['DIRECTION']],
                    'project_name' => $projectName,
                    'group_name' => ($row['GROUP'] == $groupForAbscenseId && $row['TYPE'] != 'TASK') ? $row['TITLE'] : $groups[$row['GROUP']],
                    'entity_type' => $row['TYPE'] == 'TASK' ? 'Задача' : 'Событие',
                    'title' => $row['TITLE'],
                    'date_from'=>!empty($row['DATE_FROM']) ? $row['DATE_FROM']->format('d-m-Y') : '00.00.00',
                    'date_to'=>!empty($row['DATE_TO']) ? $row['DATE_TO']->format('d-m-Y') : '00.00.00',
                    'estimate'=> $row['TIME_ESTIMATE'],
                    'elapsed'=> $row['ELAPSED'],
                    //'project_status'=> $progectStatuses[$row['CREATED_STATUS']] && $row['TYPE'] == 'TASK' ? $progectStatuses[$row['CREATED_STATUS']] : 'Н/Д',
                    'project_status' => 'Н/Д',
                    'project_client' => '',
                    'project_group_number' => '',
                    'project_curator_name' => '',
                    'project_supervisor_name' => '',
                    'project_date_plan_start' => '',
                    'project_date_plan_finish' =>'',
                    'nrv_r'  => '',//НРВ Сотрудник
                    'nrv_d'  => '',//НРВ Подразделение
                    'workload_r' => '',//% Загрузки Сотрудник.
                    'workload_d' => '',//% Загрузки Подразделение.
                    'abs_nrv_r'  => '',//НРВ Сотрудник без учета бол/отп
                    'abs_nrv_d'  => '',//НРВ Подразделение без учета бол/отп
                    'abs_workload_r' => '',//% Загрузки Сотрудник без учета бол/отп.
                    'abs_workload_d' => '',//% Загрузки Подразделение без учета бол/отп.
                    '_project_id' => $row['PROJECT_ID'],
                    '_responsible_id' => $row['RESPONSIBLE_ID'],
                    '_department_id' => $row['DEPARTMENT_ID'],
                    '_date_index' => $dateIndex,
                    
                ];
                if(!empty($row['PROJECT_ID']))
                {
                    $element['project_status'] = 'В разработке';
                }
                
                if(!empty($row['PROJECT_ID']) && !in_array($row['PROJECT_ID'], $projectIds))
                {
                    $projectIds[] = $row['PROJECT_ID'];

                    $date = new \Bitrix\Main\Type\DateTime($row['DATE_REPORT']);
                    $upd_at = $date->format('d.m.Y');
                    $history = \Fusion\Project\HistoryTable::getRow([
                        'filter' => [
                            'PROJECT_ID' => $row['PROJECT_ID'],
                            'FIELD_NAME' => 'STATUS',
                            '<=CREATE_AT' => $upd_at.' 23:59:59',
                            //'>=CREATE_AT' => $upd_at.' 00:00:00',
                        ],
                        'select' => ['ID', 'PROJECT_ID', 'CREATE_AT', 'FROM_VALUE', 'TO_VALUE'],
                        'order' => ['CREATE_AT' => 'DESC'],
                    ]);

                    if ($row['TYPE'] == 'TASK') {
                        if (is_array($history) && !empty($history['TO_VALUE'])) {
                            $element['project_status'] = $progectStatuses[$history['TO_VALUE']];
                        }
                    }

                }

                unset($history);
                $elements[] = $element;
            }

            $projects = [];
            
            if(!empty($projectIds))
            {
                $projectsFields = ProjectEntity::GetList([
                    'filter' => [
                        'ID' => $projectIds
                    ],
                    'runtime' => [
                        new \Bitrix\Main\Entity\ExpressionField('CURATOR_NAME', 'CONCAT(%s, " ", %s, " ", %s)', [
                            'CURATOR.LAST_NAME',
                            'CURATOR.NAME',
                            'CURATOR.SECOND_NAME'
                        ]),
                        new \Bitrix\Main\Entity\ExpressionField('SUPERVISOR_NAME', 'CONCAT(%s, " ", %s, " ", %s)', [
                            'SUPERVISOR.LAST_NAME',
                            'SUPERVISOR.NAME',
                            'SUPERVISOR.SECOND_NAME'
                        ]),
                    ],
                    'select' => [
                        'ID', 
                        'TITLE', 
                        'STATUS', 
                        'CLIENT_TITLE' => 'CLIENT.TITLE',
                        'GROUP_NUMBER' => 'UF_GROUP_NUMBER',
                        'CURATOR_NAME',
                        'SUPERVISOR_NAME',
                        'DATE_PLAN_START',
                        'DATE_PLAN_FINISH'
                        
                    ]
                ]);
                
                foreach($projectsFields as $projectFields){
                    $projects[$projectFields['ID']] = [
                        'id' => $projectFields['ID'],
                        'title' => $projectFields['TITLE'],
                        'status' => $progectStatuses[$projectFields['STATUS']],
                        'client' => $projectFields['CLIENT_TITLE'],
                        'group_number' => $projectFields['GROUP_NUMBER'],
                        'curator_name' => $projectFields['CURATOR_NAME'],
                        'supervisor_name' => $projectFields['SUPERVISOR_NAME'],
                        'date_plan_start' => $projectFields['DATE_PLAN_START'],
                        'date_plan_finish' => $projectFields['DATE_PLAN_FINISH'],
                    ];
                }
            }

            $metrics = $this->getMetrics($date_report_from_query, $date_report_to_query);
            $absMetrics = $this->getAbsMetrics($date_report_from_query, $date_report_to_query);

            $elements = array_map(function($e) use($projects, $metrics, $absMetrics){
               $projectId = $e['_project_id']; 
               unset($e['_project_id']);
               
               $responsibleId = $e['_responsible_id'];
               unset($e['_responsible_id']);
               
               $departmentId = $e['_department_id'];
               unset($e['_department_id']);
               
               $dateIndex = $e['_date_index'];
               unset($e['_date_index']);
               
               $target = $projects[$projectId];
               
               $responsibleMetric = (int)$metrics[$dateIndex]['responsible'][$responsibleId];
               $departmentMetric = (int)$metrics[$dateIndex]['departments'][$departmentId];
               
               $responsibleAbs = (int)$absMetrics[$dateIndex]['responsible'][$responsibleId];
               $departmentAbs = (int)$absMetrics[$dateIndex]['departments'][$departmentId];
               
               $responsibleAbsMetric = $responsibleMetric - $responsibleAbs;
               $departmentAbsMetric = $departmentMetric - $departmentAbs;

               $elapsed = $e['elapsed'] ?? 0;
               
               $e['estimate'] = !empty( $e['estimate'] ) ? $e['estimate'] / 3600 : '0';
               $e['elapsed'] = !empty( $e['elapsed'] ) ? $e['elapsed'] / 3600 : '0';
               //$e['project_status'] = $target['status'] ?? 'Н/Д';
               $e['project_client'] = $target['client'] ?? 'Н/Д';
               $e['project_group_number'] = $target['group_number'] ?? 'Н/Д';
               $e['project_curator_name'] = $target['curator_name'] ?? 'Н/Д';
               $e['project_supervisor_name'] = $target['supervisor_name'] ?? 'Н/Д';
               $e['project_date_plan_start'] = $target['date_plan_start'] ? $target['date_plan_start']->format('d-m-Y') : '00.00.00';
               $e['project_date_plan_finish'] = $target['date_plan_finish'] ? $target['date_plan_finish']->format('d-m-Y'): '00.00.00';
               $e['nrv_r'] = $responsibleMetric == 0 ? '0' : $responsibleMetric / 3600;
               $e['nrv_d'] = $departmentMetric == 0 ? '0' : $departmentMetric  / 3600;
               $e['workload_r'] = $elapsed / $responsibleMetric;
               $e['workload_d'] = $elapsed / $departmentMetric;
               $e['abs_nrv_r'] = $responsibleAbsMetric == 0 ? '0' : $responsibleAbsMetric / 3600;
               $e['abs_nrv_d'] = $departmentAbsMetric == 0 ? '0' : $departmentAbsMetric  / 3600;
               $e['abs_workload_r'] = $responsibleAbsMetric == 0 ? '0' : $elapsed / $responsibleAbsMetric;
               $e['abs_workload_d'] = $departmentAbsMetric == 0 ? '0' : $elapsed / $departmentAbsMetric;
               
               return $e;
            }, is_array($elements) ? $elements : []);

            $fileName = 'detail_report' . time() . '.xlsx';
            $filePath = sys_get_temp_dir() . '/' . $fileName;

            $writer = WriterEntityFactory::createXLSXWriter();

            $writer->openToFile($filePath);

            // Установка ширины столбцов
            $i = 1;
            $maxWidth = 20;

            foreach ($header as $code => $value) {
                $width = max(array_map('strlen', array_column($elements, $code))) + 2;

                if ($width > $maxWidth) {
                    $width = $maxWidth;
                }

                $writer->setColumnWidth($width, $i);

                $i++;
            }

            // Формирование заголовков
            $headerStyle = (new StyleBuilder())
                ->setFontBold()
                ->setShouldWrapText()
                ->setCellAlignment(CellAlignment::CENTER)
                ->build();

            $headerRow = WriterEntityFactory::createRowFromArray($header, $headerStyle);
            $writer->addRow($headerRow);

            // Формирование строк
            foreach ($elements as $element) {
                $cells = [];

                foreach ($element as $column => $cellValue) {
                    $cellStyle = (new StyleBuilder())
                        ->setShouldWrapText()
                        ->setCellAlignment(CellAlignment::RIGHT);

                    if (in_array($column, ['estimate', 'elapsed'])) {
                        $cellStyle->setFormat('0.00');
                    }

                    if (in_array($column, ['workload_r', 'workload_d'])) {
                        $cellStyle->setFormat('0.00%');
                    }

                    if (in_array($column, ['abs_workload_r', 'abs_workload_d'])) {
                        $cellStyle->setFormat('0.00%');
                    }

                    $cellStyle = $cellStyle->build();

                    $cells[] = WriterEntityFactory::createCell($cellValue, $cellStyle);
                }

                $row = WriterEntityFactory::createRow($cells);
                $writer->addRow($row);
            }

            $writer->close();

            $arFile = \CFile::MakeFileArray($filePath);
            $fileId = \CFile::SaveFile($arFile, 'timesheet');

            return \CFile::GetPath($fileId);
    }

    private function getMetrics($date_report_from_query, $date_report_to_query)
    {
        $sqlReports = "SELECT RESPONSIBLE_ID, DEPARTMENT_ID, STATS_NORM, DATE_REPORT, DATE_FORMAT(DATE_REPORT, \"%Y.%m\") as DATE_INDEX
        FROM f_ts_report report WHERE 1 {$date_report_from_query} {$date_report_to_query}";

        $connection = System::getConnection();
        $query = $connection->query($sqlReports);
        
        $metrics = [];
        
        foreach($query as $row)
        {
            $index = $row['DATE_INDEX'];
            
            if(empty($index))
            {
                continue;
            }
            
            // Собираем суммарно нормы ответственных сотрудников по каждому месяцу
            if(empty($metrics[$index]['responsible'][$row['RESPONSIBLE_ID']]))
            {
                $metrics[$index]['responsible'][$row['RESPONSIBLE_ID']] = $row['STATS_NORM'];
            }
            else
            {
                $metrics[$index]['responsible'][$row['RESPONSIBLE_ID']] += $row['STATS_NORM'];
            }
            
            // Собираем суммарно нормы подразделений по каждому месяцу
            if(empty($metrics[$index]['departments'][$row['DEPARTMENT_ID']]))
            {
                $metrics[$index]['departments'][$row['DEPARTMENT_ID']] = $row['STATS_NORM'];
            }
            else
            {
                $metrics[$index]['departments'][$row['DEPARTMENT_ID']] += $row['STATS_NORM'];
            }
        }
        
        return $metrics;
    }
    
    private function getAbsMetrics($date_report_from_query, $date_report_to_query)
    {
        $sqlReports = "SELECT DATE_FORMAT(DATE_REPORT, \"%Y.%m\") as DATE_INDEX, report.`RESPONSIBLE_ID`, report.`DEPARTMENT_ID`, report.`STATS_NORM`
            FROM f_ts_absence tsa
            LEFT JOIN f_ts_report report
            ON report.ID = tsa.REPORT_ID
            WHERE tsa.TYPE IN('LEAVESICK', 'VACATION') {$date_report_from_query} {$date_report_to_query}";
        
        $connection = System::getConnection();
        $query = $connection->query($sqlReports);
        
        $metrics = [];
        
        foreach($query as $row)
        {
            $index = $row['DATE_INDEX'];
            
            if(empty($index))
            {
                continue;
            }
            
            // Собираем суммарно нормы ответственных сотрудников по каждому месяцу
            if(empty($metrics[$index]['responsible'][$row['RESPONSIBLE_ID']]))
            {
                $metrics[$index]['responsible'][$row['RESPONSIBLE_ID']] = $row['STATS_NORM'];
            }
            else
            {
                $metrics[$index]['responsible'][$row['RESPONSIBLE_ID']] += $row['STATS_NORM'];
            }
            
            // Собираем суммарно нормы подразделений по каждому месяцу
            if(empty($metrics[$index]['departments'][$row['DEPARTMENT_ID']]))
            {
                $metrics[$index]['departments'][$row['DEPARTMENT_ID']] = $row['STATS_NORM'];
            }
            else
            {
                $metrics[$index]['departments'][$row['DEPARTMENT_ID']] += $row['STATS_NORM'];
            }
        }
        
        return $metrics;
    }
    
    public function actualizeReportsAction($url)
    {
        if (! \Fusion\Timesheet\Permission::isAdmin() && ! \Fusion\Timesheet\Permission::isPersonnelDepartment()) {
            throw new \Exception('Недостаточно прав для совершения действия');
        }
        
        $response = parse_url($url);
        
        if (! mb_parse_str($response['query'], $query_params)) {
            throw new \Exception('Ошибка обработки параметров');
        }
        
        global $APPLICATION;
        
        ob_start();
        
        $data = $APPLICATION->IncludeComponent('bitrix:intranet.absence.calendar', '', array(
            "AJAX_CALL" => "DATA",
            "CALLBACK" => 'DATA_RETURN',
            'SITE_ID' => $query_params['site_id'],
            'IBLOCK_ID' => $query_params['iblock_id'],
            'CALENDAR_IBLOCK_ID' => $query_params['calendar_iblock_id'],
            "FILTER_SECTION_CURONLY" => $query_params['section_flag'] == 'Y' ? 'Y' : 'N',
            "TS_START" => $query_params['TS_START'],
            "TS_FINISH" => $query_params['TS_FINISH'],
            'PAGE_NUMBER' => $query_params['PAGE_NUMBER'],
            "SHORT_EVENTS" => $query_params['SHORT_EVENTS'],
            "USERS_ALL" => $query_params['USERS_ALL'],
            "CURRENT_DATA_ID" => $query_params['current_data_id']
        ));
        
        $prepared_data = array_filter(array_map(function ($r) {
            $part = [
                'responsible_id' => $r['ID'],
                'events' => array_filter(array_map(function ($ed) {
                    return [
                        'NAME' => $ed['NAME'],
                        'DATE_FROM' => $ed['DATE_FROM'],
                        'DATE_TO' => $ed['DATE_TO'],
                        'TYPE' => $ed['TYPE'],
                        'DT_FROM_TS' => strtotime($ed['DATE_FROM'] . ' 00:00:00'),
                        'DT_TO_TS' => strtotime($ed['DATE_TO'] . ' 23:59:59')
                    ];
                }, is_array($r['DATA']) ? $r['DATA'] : []), function ($e) {
                    return true;
                })
            ];
            
            return $part;
        }, is_array($data) ? $data : []), function ($user_data) {
            return ! empty($user_data['events']);
        });
        
        $from = new \DateTime(date('d.m.Y', $query_params['TS_START']));
        $to = new \DateTime(date('d.m.Y 00:00:1', $query_params['TS_FINISH']));
        
        $period = new \DatePeriod($from, new \DateInterval('P1D'), $to);
        
        $list_dates = array_map(function ($item) {
            return $item->format('d.m.Y');
        }, iterator_to_array($period));

        $actual_events = [];
        
        $group_id = \Fusion\Timesheet\Permission::getGroupIdByCode(\Fusion\Timesheet\Permission::EMPLOYE_GROUP_CODE);
        $users = \CGroup::GetGroupUser($group_id);

        // Task: 45406
        $res = \Bitrix\Main\UserTable::getList([
            'select' => ['ID'],
            'filter' => [
                'ACTIVE' => false,
                '!WORK_COMPANY' => false,                
            ],
        ])->fetchAll();
        $terminateds = array_column($res, 'ID');
        if(count($terminateds)){
            $users = array_unique (array_merge ($users, $terminateds));
        }

        $userShedules = [];
        
        foreach ($list_dates as $one_day) {
            
            $time = strtotime($one_day);
            
            foreach ($prepared_data as $user_data) {
                
                $responsible_id = $user_data['responsible_id'];
                
                if (in_array($responsible_id, $users)) {
                    
                    if (empty($userShedules[$responsible_id])) {
                        $userShedules[$responsible_id] = new Timeman\UserShedule($responsible_id);
                    }
                    
                    $userShedule = new Timeman\UserShedule($responsible_id);
                    
                    $dateTime = new \DateTime($one_day);
                    
                    if ($userShedule->isWorkDate($dateTime)) {
                        foreach ($user_data['events'] as $event) {
                            if ($time >= $event['DT_FROM_TS'] && $time <= $event['DT_TO_TS']) {
                                $actual_events[$one_day][] = [
                                    'responsible_id' => $responsible_id,
                                    'type' => $event['TYPE'],
                                    'name' => $event['NAME']
                                ];
                            }
                        }
                    }
                }
            }
        }
        
        $APPLICATION->restartBuffer();
        
        $reports = [];
        
        foreach ($actual_events as $date => $events) {
            foreach ($events as $event) {
                $objectDate = new \Bitrix\Main\Type\Date($date);
                $user_id = $event['responsible_id'];
                
                $params = [
                    'TITLE' => $event['name'],
                    'TYPE' => $event['type'],
                    'DATE_ABSENCE' => $objectDate
                ];
                
                $builder = new \Fusion\Timesheet\Builder($user_id, $objectDate);
                
                if ($event['type'] == 'PERSONAL') {
                    if ($builder->isReportExists()) {
                        $id = $builder->getCurrentReportId();
                        TimeSheet::delete($id);
                        
                        $builder = new \Fusion\Timesheet\Builder($user_id, $objectDate);
                    }
                    
                    $builder->build();
                    
                    $report_id = $builder->getCurrentReportId();
                } else {
                    $report_id = $builder->buildAbsenceReport($params);
                }
                
                $reports[] = $report_id;
            }
        }
        
        $from = new \Bitrix\Main\Type\Date(date('d.m.Y', $query_params['TS_START']));
        $to = new \Bitrix\Main\Type\Date(date('d.m.Y', $query_params['TS_FINISH']));
        
        $reports_to_draft = TimeSheet::getList([
            'filter' => [
                '!ID' => $reports,
                'STATUS' => [
                    Dictionaries::ABSENCE_STATUS,
                    Dictionaries::SUBMITTED_AUTO_STATUS
                ],
                '>=DATE_REPORT' => $from,
                '<=DATE_REPORT' => $to
            ],
            'select' => [
                'ID',
                'RESPONSIBLE_ID',
                'DATE_REPORT'
            ]
        ]);
        
        foreach ($reports_to_draft as $report_to_draft) {
            $responsible_id = $report_to_draft['RESPONSIBLE_ID'];
            $report_date = $report_to_draft['DATE_REPORT'];
            
            $builder = new \Fusion\Timesheet\Builder($responsible_id, $report_date);
            
            $builder->cleanReport(true);
            $builder->actualize(true);
        }
        
        return true;
    }
    
    /**
     * Получение параметров для построения фильтра
     *
     * @param array $option
     * @return array
     */
    private function getFilterFields($option = []): array
    {
        $filterFields = [
            [
                'id' => 'DATE_REPORT',
                'name' => 'Дата ТШ',
                'type' => 'date',
                'default' => true
            ],
            [
                'id' => 'STATUS',
                'name' => 'Статус',
                'type' => 'list',
                'params' => [
                    'multiple' => 'Y'
                ],
                'items' => [] + Dictionaries::getStatuses(),
                'default' => true
            ],
            [
                'id' => 'RESPONSIBLE_ID',
                'name' => 'ФИО сотрудника',
                'type' => 'entity_selector',
                'params' => [
                    'multiple' => true,
                    'dialogOptions' => [
                        'height'       => 200,
                        'entities'     => [
                            [
                                'id' => 'user',
                                'options' => [
                                    'inviteEmployeeLink' => false,
                                    'intranetUsersOnly'  => true,
                                ]
                            ],
                            [
                                'id' => 'fired-user',
                            ],
                        ],
                        'showAvatars'  => true,
                        'dropdownMode' => false,
                    ]
                ],
                'default' => true
            ],
            [
                'id' => 'DEPARTMENT_ID',
                'name' => 'Подразделение',
                'type' => 'dest_selector',
                'params' => [
                    'apiVersion' => 3,
                    'contextCode' => 'CRM',
                    'useClientDatabase' => 'N',
                    'enableAll' => 'N',
                    'enableDepartments' => 'Y',
                    'enableUsers' => 'N',
                    'enableSonetgroups' => 'N',
                    'allowEmailInvitation' => 'N',
                    'allowSearchEmailUsers' => 'N',
                    'departmentSelectDisable' => 'N',
                    'enableCrm' => 'N',
                    'enableCrmContacts' => 'N',
                    'convertJson' => 'N',
                    'multiple' => false
                ],
                'default' => true
            ]
        ];
        
        return $filterFields;
    }
    
    /**
     * Получение параметров для построения списка
     *
     * @param array $option
     * @return array
     */
    private function getColumns($option = []): array
    {
        $columns = [
            [
                'id' => 'RESPONSIBLE_ID',
                'name' => 'Сотрудник',
                'sort' => 'RESPONSIBLE_ID',
                'default' => true
            ],
            [
                'id' => 'DEPARTMENT_NAME',
                'name' => 'Подразделение',
                'sort' => 'DEPARTMENT_ID',
                'default' => true
            ],
            [
                'id' => 'DATE_REPORT',
                'name' => 'Дата ТШ',
                'sort' => 'DATE_REPORT',
                'default' => true
            ],
            [
                'id' => 'STATUS',
                'name' => 'Статус',
                'sort' => 'STATUS',
                'default' => true
            ],
            [
                'id' => 'STATS_TOTAL',
                'name' => 'Итого',
                'sort' => 'STATS_TOTAL',
                'default' => true
            ],
            [
                'id' => 'STATS_TASK',
                'name' => 'Время по задачам',
                'sort' => 'STATS_TASK',
                'default' => true
            ],
            [
                'id' => 'STATS_EVENT',
                'name' => 'Время по событиям',
                'sort' => 'STATS_EVENT',
                'default' => true
            ],
            [
                'id' => 'STATS_NORM',
                'name' => 'Норма',
                'sort' => 'STATS_NORM',
                'default' => false
            ]
        ];
        
        return $columns;
    }
    
    /**
     * Получение массива для фильтрации данных
     *
     * @param string $grid_id
     * @return array
     */
    private function prepairFilter($grid_id, $grid_filter): array
    {
        $filter = [];
        
        $filterOption = new \Bitrix\Main\UI\Filter\Options($grid_id, $this->getFilterDefaultPesets());
        
        $filterData = $filterOption->getFilter([]);
        
        foreach ($filterData as $k => $v) {
            
            if (in_array($k, [
                'RESPONSIBLE_ID',
                'DEPARTMENT_ID'
            ])) {
                $v = str_replace('CRMCOMPANY', '', $v);
                $v = str_replace('U', '', $v);
                $v = str_replace('DR', '', $v);
            }
            
            $filter[$k] = $v;
        }
        
        $filterPrepared = \Bitrix\Main\UI\Filter\Type::getLogicFilter($filter, $grid_filter);
        
        if (! empty($filter['STATUS'])) {
            $filterPrepared['STATUS'] = $filter['STATUS'];
        }
        
        return $filterPrepared;
    }
    
    /**
     * Получение отображения пользователя
     *
     * @param integer $id
     * @return string
     */
    private function getUserViewHtml($id): string
    {
        if (empty($id)) {
            return '';
        }
        
        global $APPLICATION;
        
        ob_start();
        
        $APPLICATION->IncludeComponent("bitrix:main.user.link", "", Array(
            "CACHE_TYPE" => "N",
            "CACHE_TIME" => "7200",
            "ID" => $id,
            "NAME_TEMPLATE" => "#NOBR##LAST_NAME# #NAME##/NOBR#",
            "SHOW_LOGIN" => "Y",
            "USE_THUMBNAIL_LIST" => "Y"
            ));
        
        return ob_get_clean();
    }
    
    public function getFilterDefaultPesets()
    {
        return [
            'submitted_in_month' => [
                'name' => 'Сданы в текущ. месяце',
                'default' => true,
                'fields' => [
                    'DATE_REPORT_datesel' => \Bitrix\Main\UI\Filter\DateType::CURRENT_MONTH,
                    'DATE_REPORT_from' => '',
                    'DATE_REPORT_to' => '',
                    'DATE_REPORT_days' => '',
                    'DATE_REPORT_month' => '',
                    'DATE_REPORT_quarter' => '',
                    'DATE_REPORT_year' => '',
                    'STATUS' => [
                        'SUBMITTED',
                        'SUBMITTED_AUTO',
                        'ABSENCE'
                    ]
                ]
            ]
        ];
    }
    
    public static function getTimeStringBySeconds($seconds)
    {
        $hours = floor($seconds / 3600);
        $minutes = round(($seconds / 3600 - $hours) * 60);
        
        $hours = $hours < 10 ? "0{$hours}" : $hours;
        $minutes = $minutes < 10 ? "0{$minutes}" : $minutes;
        
        return "{$hours} ч {$minutes} м";
    }
    
    /**
     * Расширение фильтра с учётом роли пользователя
     *
     * @param array $filter
     * @return array
     */
    public function getPermissionFilter($filter): array
    {
        if (! Permission::isAdmin() && ! Permission::isPed() && ! Permission::isObserver()) {
            $subordinateUsers = [
                System::getUser()->getid()
            ];
            
            $subordinateUsersCollection = \CIntranetUtils::GetSubordinateEmployees($USER_ID = null, $bRecursive = true, $onlyActive = 'Y');
            
            while ($subordinateUser = $subordinateUsersCollection->fetch()) {
                $subordinateUsers[] = $subordinateUser['ID'];
            }
            
            $filter[] = [
                'LOGIC' => 'AND',
                'RESPONSIBLE_ID' => $subordinateUsers
            ];
        }
        
        return $filter;
    }
}
