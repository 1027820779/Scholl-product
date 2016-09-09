<?php
Class ExcelAction extends CommonAction{

    //显示
    public function index()
    {
        //本年
        //date_default_timezone_set('Asia/Urumqi');
        $year=date('Y');
        for($i=0;$i<20;$i++){

            $yr[]=array(
                'l'=>$year-$i,
                'b'=>$year-$i+1
            );
        }
        $this->sch_year=$yr;

        if(IS_POST){
            //学年
            $school_year=$_POST['school_year'];
            //学期
            $term=$_POST['term'];
            $user = M('student_info')->where(array('school_year' => $school_year,'term'=>$term))->select();
            for($i = 0; $i < count($user); $i++) {
                if($user[$i]['class'] != $user[$i-1]['class'] && $user[$i]['class'] != $user[$i+1]['class'] ){
                    $cl[]=array(
                        'class'=>$user[$i]['class'],
                        'school_year'=>$user[$i]['school_year'],
                        'term'=>$user[$i]['term']
                    );
                }

            }
            //一共有多少个班,属于那个学年，那个学期
            $this->class=$cl;
        }else{
            $filed=array('class','school_year','term');
            $user = M('student_info')->field($filed)->select();
        }

        //echo(count($user));
        for($i = 0; $i < count($user); $i++) {
            if($user[$i]['class'] != $user[$i-1]['class']){
                $cl[]=array(
                    'class'=>$user[$i]['class'],
                    'school_year'=>$user[$i]['school_year'],
                    'term'=>$user[$i]['term']
                );
            }

        }
        //一共有多少个班,属于那个学年，那个学期
        $this->class=$cl;


        $this->display();
    }

    //删除班级成绩
    public function delete(){
        //学年
        $school_year=$_POST['school_year'];
        //学期
        $term=$_POST['term'];
        //班级
        $class=$_POST['class'];

        if(M('student_info')->where(array('class'=>$class,'school_year'=>$school_year,'term'=>$term))->delete()){
            $this->success('删除成功',U('Admin/Excel/index'));
        }else{
            $this->error('删除失败');
        }

    }

    //上传
    public  function upload(){
        import('ORG.Net.UploadFile');
        $upload=new UploadFile();


        $upload->maxSize=1024*1024*500;   // 上传文件的最大值
        $upload->allowExts=array('xls','xlsx');    // 允许上传的文件后缀 留空不作后缀检查
        $upload->autoSub=true;// 启用子目录保存文件
        $upload->subType = 'date';
        $upload->dateFormat='Ymd';//时间格式
        $upload->savePath='./uploads/';   // 上传文件保存路径



        if (!$upload->upload()) {
            //如果上传失败以下
            echo("<script>alert('不能上传过大的文件')</script>");
            $this->error();
        }else{
            $info=$upload->getUploadFileInfo();
            $exts=$info[0]['extension'];
            $file_url=$info[0]['savepath'].$info[0]['savename'];
            $name_s=$info[0]['name'];
            $this->read_exsl($file_url,$exts,$name_s);
        }


    }
    //读取exel
    public function read_exsl($file_url,$exts,$name_s){
        $frefix=substr($name_s, strrpos($name_s, '.'));
        $na=trim($name_s,$frefix);
        import('ORG.Util.PHPExcel');
        $php_excel=new PHPExcel();

        if($exts=='xls'){
            import('ORG.Util.PHPExcel.Reader.Excel5');
            $php_reader=new PHPExcel_Reader_Excel5();
        }else if($exts=='xlsx'){
            import('ORG.Util.PHPExcel.Reader.Excel2007');
            $php_reader=new PHPExcel_Reader_Excel2007();
        }

        $PHPExcel=$php_reader->load($file_url);

        $current_sheet=$PHPExcel->getSheet(0);
        $all_colmun=$current_sheet->getHighestColumn();
        $all_row=$current_sheet->getHighestRow();
        for($currentRow=1;$currentRow<=$all_row;$currentRow++){
            for($currentColmun='A';$currentColmun<=$all_colmun;$currentColmun++){
                $address=$currentColmun.$currentRow;
                $data[$currentRow][$currentColmun]=$current_sheet->getCell($address)->getValue();
            }
        }
        $this->save_import($data,$na);
    }

    //把exel数据保存到数据库
    public function save_import($data,$na){
        M('result_db')->where(array('id'=>array('gt',0)))->delete();

        foreach($data as $k=>$v){
            if($k>=1){
                $dat=array(
                    'student_number'=>$v['B'],
                    'test_number'=>$v['C'],
                    'class'=>$v['D'],
                    'student_name'=>$v['E'],
                    'math_num'=>$v['F'],
                    'language_num'=>$v['G'],
                    'chines_num'=>$v['H'],
                    'politics_num'=>$v['I'],
                    'history_num'=>$v['J'],
                    'geography_num'=>$v['K'],
                    'biology_num'=>$v['L'],
                    'chemistry_num'=>$v['M'],
                    'physics_num'=>$v['N'],
                    'eligible_subjects'=>$v['O'],
                    'total_points'=>$v['P'],
                    'average'=>$v['Q'],
                    'class_level'=>$v['R'],
                    'school_level'=>$v['S'],
                );
            }
            M('result_db')->add($dat);
        }

        $a=M('result_db')->field('id',true)->select();

        //所有的栏目
        $this->subject=$a[0];
        //本年
        //date_default_timezone_set('Asia/Urumqi');
        $year=date('Y');
        for($i=0;$i<20;$i++){

            $yr[]=array(
                'l'=>$year-$i,
                'b'=>$year-$i+1
            );
        }
        $this->sch_year=$yr;
        //去掉扩张名的文件名
        $this->f_name=$na;

        $this->display('save_import');

    }

    //选择科目
    public  function save_imp(){
        //所选的科目
        $subjects=$_POST;

        $filed=array('student_number','test_number','class','student_name');
        //全部数据
        $all_dat=M('result_db')->field($filed)->select();

                foreach($all_dat as $k=>$v){
            if($k>=1){
                $info[]=array(
                    'student_number'=>$v['student_number'],
                    'test_number'=>$v['test_number'],
                    'class'=>$v['class'],
                    'student_name'=>$v['student_name']
                );
            }

        }

        //学年
        $school_year=$subjects['sch_year'];

        //学期
        $term=$subjects['term'];


        //所选的科目
        $sub=$subjects['subjects'];
        $all_subject=M('result_db')->field($sub)->select();



        //科目数
        $sub_length=count($all_subject[0]);
        //除去素组第一个
        $all_subject1=array_splice($all_subject,1);
        //成绩总行数
        $all_sub_len=count($all_subject1);
        //名单总数
        $all_info_len=count($info);

        $sub_n=$all_subject[0];

        for($i = 0; $i <= count($info); $i++) {

            $ar=$all_subject1[$i];
            for ($j = 0; $j <= (count($ar)-1); $j++) {

                $ll=array(
                    'subject'=>$sub_n[$sub[$j]],
                    'result'=>$all_subject1[$i][$sub[$j]],
                    'school_year'=>$school_year,
                    'term'=>$term
                );
                $in=$info[$i];
                $array3[]=array_merge($in,$ll);
            }

        }

        if(M('student_info')->addAll($array3)){
            $this->success('添加成功',U('Admin/Excel/see'));
        }else{
            $this->error('添加失败');
        }

    }


    //查看
    public function see(){
        $user = M('student_info')->field('class')->select();
        $a=array_count_values($user);

        //echo(count($user));
        for($i = 0; $i < count($user); $i++) {
            if($user[$i]['class'] != $user[$i-1]['class']){
                $cl[]=array(
                    'class'=>$user[$i]['class'],
                );
            }

        }
        //一共有多少个班
        $this->class=$cl;

        //本年
        //date_default_timezone_set('Asia/Urumqi');
        $year=date('Y');
        for($i=0;$i<10;$i++){

            $yr[]=array(
                'l'=>$year-$i,
                'b'=>$year-$i+1
            );
        }
        $this->sch_year=$yr;

        $this->display();
    }

    public function deal(){
            $year=date('Y');
            for($i=0;$i<10;$i++){

                $yr[]=array(
                    'l'=>$year-$i,
                    'b'=>$year-$i+1
                );
            }
            $this->sch_year=$yr;

            //学年
            $school_year=$_POST['school_year'];
            $this->s_y=$school_year;

            //学期
            $term=$_POST['term'];
            $this->term=$term;

            //班级
            $class=$_POST['class'];
            $this->class=$class;

        //按照班级信息
        $user = M('student_info')->where(array('school_year' => $school_year,'term'=>$term,'class'=>$class))->select();
            for($i = 0; $i < count($user); $i++) {
                if ($user[$i]['test_number'] != $user[$i - 1]['test_number']) {
                    $all_name[] = array(
                        'student_name' => $user[$i]['student_name'],
                    );
                }
            }
        //按照班级信息
        for($i = 0; $i < count($user)/count($all_name); $i++) {
            $all_sub_name[]=array(
                'subject'=>$user[$i]['subject'],
            );
        }

        //按照全校信息
        $all_info = M('student_info')->where(array('school_year' => $school_year,'term'=>$term))->select();
        for($i = 0; $i < count($all_info); $i++) {
            if ($all_info[$i]['test_number'] != $all_info[$i - 1]['test_number']) {
                $all_test_num[] = array(
                    'test_number' => $all_info[$i]['test_number'],
                );
            }
        }



        //按照全校信息
        //全校平均分
        $ping=array();
        $student_info=M();
        for($i = 0; $i < count($all_test_num); $i++) {
            $i_t=$all_test_num[$i]['test_number'];
            $avv=$student_info->query("select avg(result) as a from hd_student_info where test_number=$i_t");
            $avv1=number_format($avv[0][a],2);
            array_push($ping,$avv1);
        }



        //所有的科目名称
        $this->all_s_n=$all_sub_name;

        //成绩集合包
        $oku=array();

        for ($i = 0; $i < count($user); $i++) {
            $all_sl[]= array($user[$i]['subject']=>$user[$i]['result']);
           array_push($all_su,$all_sl);
            if (($i + 1) % count($all_sub_name) == 0) {
               array_push($oku,$all_sl);
                $all_sl =[];
            }
        }




        for($i = 0; $i < count($user); $i++) {
            if($user[$i]['test_number'] != $user[$i-1]['test_number']){
                //基本信息集合
                $cc[]=array(
                    'student_number'=>$user[$i]['student_number'],
                    'test_number'=>$user[$i]['test_number'],
                    'class'=>$user[$i]['class'],
                    'student_name'=>$user[$i]['student_name']
                );
                //学期和学年集合
                $two_info[]=array(
                    'school_year'=>$user[$i]['school_year'],
                    'term'=>$user[$i]['term']
                );
            }

        }
        //average|平均分; total_points|总分;standard|合格科目书;class_order|班级等级;

        //集合总成绩
        $lo=array();

        $rtt=0;
        $h=0;
        for($i = 0; $i <= (count($cc)-1); $i++) {
            $arr=array();
            $total_points=array();
            for ($j = 0; $j <= (count($all_sub_name)-1); $j++) {
                $result = $oku[$i][$j][$all_sub_name[$j]['subject']];
                if ($result >= 60) {
                    $h++;
                }
                $rtt+=$result;
                $total_points['total_points']=$rtt;
                array_push($arr,$result);
            }
            $average=array();
            $average['average']=number_format($rtt/count($all_sub_name),2);
            $standard=array();
            $standard['standard']=$h;
            $m=array_merge($cc[$i],$arr,$average,$total_points,$standard,$two_info[$i]);
            array_push($lo,$m);
            $arr=[];
            $rtt=0;
            $h=0;
        }

        //按照班级
        //大到小排序
        $z_po=array();
        for($i = 0; $i <= (count($lo)-1); $i++) {
            $rr=$lo[$i]['total_points'];
                array_push($z_po,$rr);
        }
        $ayu=rsort($z_po);

        $z_poo=array();
        for($i = 0; $i <= (count($lo)-1); $i++) {
            $rr=$lo[$i]['total_points'];
            array_push($z_poo,$rr);
        }

        //平均分结合包
        $z_pooo=array();
        for($i = 0; $i <= (count($lo)-1); $i++) {
            $rr=$lo[$i]['average'];
            array_push($z_pooo,$rr);
        }

        //全校等级
        //大到小排序
        $ayu1=rsort($ping);
        $loo=array();
        //班级等级
        for($i = 0; $i <= (count($lo)-1); $i++) {
            $m_z=$z_poo[$i];
            $p_z=$z_pooo[$i];
            $d=array_search($m_z,$z_po);
            $d=$d+1;
            $k=array_search($p_z,$ping);
            $k=$k+1;

            //等级集合
            $sort_n=array();
            $sort_n['class_order']=$d;
            $sort_n['schol_order']=$k;
            $nk=array_merge($lo[$i],$sort_n);
            array_push($loo,$nk);
            $sort_n=[];
        }
            $this->all_info=$loo;

            $this->display();
    }

    //导出Excel
    public function down(){
        import('Class.Excel',APP_PATH);
        $excel=new Excel();


        //学年
        $school_year=$_POST['school_year'];

        //学期
        $term=$_POST['term'];

        //班级
        $class=$_POST['class'];


        //按照班级信息
        $user = M('student_info')->where(array('school_year' => $school_year,'term'=>$term,'class'=>$class))->select();
        for($i = 0; $i < count($user); $i++) {
            if ($user[$i]['test_number'] != $user[$i - 1]['test_number']) {
                $all_name[] = array(
                    'student_name' => $user[$i]['student_name'],
                );
            }
        }
        //按照班级信息
        for($i = 0; $i < count($user)/count($all_name); $i++) {
            $all_sub_name[]=array(
                'subject'=>$user[$i]['subject'],
            );
        }

        //按照全校信息
        $all_info = M('student_info')->where(array('school_year' => $school_year,'term'=>$term))->select();
        for($i = 0; $i < count($all_info); $i++) {
            if ($all_info[$i]['test_number'] != $all_info[$i - 1]['test_number']) {
                $all_test_num[] = array(
                    'test_number' => $all_info[$i]['test_number'],
                );
            }
        }



        //按照全校信息
        //全校平均分
        $ping=array();
        $student_info=M();
        for($i = 0; $i < count($all_test_num); $i++) {
            $i_t=$all_test_num[$i]['test_number'];
            $avv=$student_info->query("select avg(result) as a from hd_student_info where test_number=$i_t");
            $avv1=number_format($avv[0][a],2);
            array_push($ping,$avv1);
        }



        //所有的科目名称
        $this->all_s_n=$all_sub_name;

        //成绩集合包
        $oku=array();

        for ($i = 0; $i < count($user); $i++) {
            $all_sl[]= array($user[$i]['subject']=>$user[$i]['result']);
            array_push($all_su,$all_sl);
            if (($i + 1) % count($all_sub_name) == 0) {
                array_push($oku,$all_sl);
                $all_sl =[];
            }
        }




        for($i = 0; $i < count($user); $i++) {
            if($user[$i]['test_number'] != $user[$i-1]['test_number']){
                //基本信息集合
                $cc[]=array(
                    'student_number'=>$user[$i]['student_number'],
                    'test_number'=>$user[$i]['test_number'],
                    'class'=>$user[$i]['class'],
                    'student_name'=>$user[$i]['student_name']
                );
                //学期和学年集合
                $two_info[]=array(
                    'school_year'=>$user[$i]['school_year'],
                    'term'=>$user[$i]['term']
                );
            }

        }

        //结合表哥第一行
                $pan=array();
                for($i=0;$i<count($all_sub_name);$i++){
                    array_push($pan,$all_sub_name[$i]['subject']);
                }
                $ccm=array(
                    'student_number'=>'学籍号',
                    'test_number'=>'考号',
                    'class'=>'班级',
                    'student_name'=>'姓名'
                );
                $ccn=array(
                    'average'=>'平均分',
                    'total_points'=>'总分',
                    'standard'=>'合格科目数',
                    'school_year'=>'学年',
                    'term'=>'学期',
                    'class_order'=>'班级等级',
                    'schol_order'=>'学校等级',
                );
        $fir=array_merge($ccm,$pan);
        $fir_1=array_merge($fir,$ccn);


        //average|平均分; total_points|总分;standard|合格科目书;class_order|班级等级;

        //集合总成绩
        $lo=array();

        $rtt=0;
        $h=0;
        for($i = 0; $i <= (count($cc)-1); $i++) {
            $arr=array();
            $total_points=array();
            for ($j = 0; $j <= (count($all_sub_name)-1); $j++) {
                $result = $oku[$i][$j][$all_sub_name[$j]['subject']];
                if ($result >= 60) {
                    $h++;
                }
                $rtt+=$result;
                $total_points['total_points']=$rtt;
                array_push($arr,$result);
            }
            $average=array();
            $average['average']=number_format($rtt/count($all_sub_name),2);
            $standard=array();
//            if($h==0){
//            $h='\'0';
//            }
            $standard['standard']=$h;
            //$standard['standard']=int();
            $m=array_merge($cc[$i],$arr,$average,$total_points,$standard,$two_info[$i]);
            array_push($lo,$m);
            $arr=[];
            $rtt=0;
            $h=0;
        }

        //按照班级
        //大到小排序
        $z_po=array();
        for($i = 0; $i <= (count($lo)-1); $i++) {
            $rr=$lo[$i]['total_points'];
            array_push($z_po,$rr);
        }
        $ayu=rsort($z_po);

        $z_poo=array();
        for($i = 0; $i <= (count($lo)-1); $i++) {
            $rr=$lo[$i]['total_points'];
            array_push($z_poo,$rr);
        }

        //平均分结合包
        $z_pooo=array();
        for($i = 0; $i <= (count($lo)-1); $i++) {
            $rr=$lo[$i]['average'];
            array_push($z_pooo,$rr);
        }

        //全校等级
        //大到小排序
        $ayu1=rsort($ping);
        $loo=array();
        //班级等级
        for($i = 0; $i <= (count($lo)-1); $i++) {
            $m_z=$z_poo[$i];
            $p_z=$z_pooo[$i];
            $d=array_search($m_z,$z_po);
            $d=$d+1;
            $k=array_search($p_z,$ping);
            $k=$k+1;

            //等级集合
            $sort_n=array();
            $sort_n['class_order']=$d;
            $sort_n['schol_order']=$k;
            $nk=array_merge($lo[$i],$sort_n);
            array_push($loo,$nk);
            $sort_n=[];
        }


        $file_name=$class.$school_year.'学年第'.$term.'学期成绩表';
        $mm=array(
            'student_number'=>$file_name
        );
//        第一行和其他信息结合
        array_unshift($loo,$fir_1);
        array_unshift($loo,$mm);
        //p($loo);



        $data=$loo;
        $excel->download($data,$file_name);
    }


    //成绩编辑
    public function edit(){
        //考号
        $test_number=I('get.test_number');

        //平均分
        $average=I('get.average');
        $this->average=$average;

        //总分
        $total_points=I('get.total_points');
        $this->total_points=$total_points;

        //合格科目数
        $standard=I('get.standard');
        $this->standard=$standard;

        //班级等级
        $class_order=I('get.class_order');
        $this->class_order=$class_order;

        //全校等级
        $schol_order=I('get.schol_order');
        $this->schol_order=$schol_order;



        $user = M('student_info')->where(array('test_number' => $test_number))->select();
        //echo(count($user));
        for($i = 0; $i < count($user); $i++) {
            $ll[]=array(
                'subject'=>$user[$i]['subject'],
                'result'=>$user[$i]['result']
            );
        }

        //各科目名称和得分数
        $this->sub_res=$ll;

        $this->nati=$user[0];
        $this->display();
    }

    //处理成绩编辑表单
    public function editHandle(){
        //考号
        $test_number=$_POST['test_number'];

        //成绩包
        $results=$_POST;

        $user = M('student_info')->where(array('test_number' => $test_number))->select();

        for($i = 0; $i < count($user); $i++) {
            $ll[]=array(
                'id'=>$user[$i]['id'],
                'subject'=>$user[$i]['subject'],
                'result'=>$user[$i]['result']
            );
        }

        for($i=0;$i<count($ll);$i++){
            //科目名称
            $s=$ll[$i]['subject'];
            //得分数
            $n=$results[$s];
            //id
            $id=$ll[$i]['id'];

            $re['result']=$n;
            M('student_info')->where(array('id'=>$id))->save($re);
        }

        $this->success('保存成功',U('see'));

    }


}
    ?>