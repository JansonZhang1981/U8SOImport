Imports UFIDA.U8.MomServiceCommon
Imports UFIDA.U8.U8MOMAPIFramework
Imports UFIDA.U8.U8APIFramework
Imports UFIDA.U8.U8APIFramework.Meta
Imports UFIDA.U8.U8APIFramework.Parameter
Imports MSXML2
Imports System.Data
Imports System.Data.OleDb
Imports System.Threading
Public Class Form1
    Public excConn As OleDbConnection

    '    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '        On Error GoTo ErrHandler
    '        Dim u8EnvCtx As New U8EnvContext
    '        u8EnvCtx.U8Login = u8login
    '        '设置上下文参数
    '        u8EnvCtx.SetApiContext("VoucherType", 12)  '上下文数据类型：int，含义：单据类型：12

    '        '第三步：构造ApiBroker对象,调用Connect,传入Api的地址标识(Url)，传入上下文
    '        Dim u8apiBroker As New U8ApiComBroker

    '        u8apiBroker.Connect("U8API/SaleOrder/Save", u8EnvCtx)
    '        '方法二是构造BusinessObject对象，具体方法如下：

    '        Dim domHead As BusinessObject
    '        domHead = u8apiBroker.GetBoParam("domHead")
    '        domHead.RowCount = 1 '设置BO对象(表头)行数，只能为一行
    '        '给BO对象(表头)的字段赋值，值可以是真实类型，也可以是无类型字符串

    '        '****************************** 以下是必输字段 *****************************
    '        domHead(0).SetValue("id", "100000002")   '主关键字段，Integer类型
    '        domHead(0).SetValue("csocode", "200000002")   '订 单 号，String类型
    '        domHead(0).SetValue("ddate", "2016-4-11")   '订单日期，Date类型
    '        domHead(0).SetValue("cbustype", "普通销售")   '业务类型，Integer类型
    '        domHead(0).SetValue("cstname", "普通销售")   '销售类型，String类型
    '        domHead(0).SetValue("ccusabbname", "商用车")   '客户简称，String类型
    '        domHead(0).SetValue("cdepname", "市场部")   '销售部门，String类型
    '        domHead(0).SetValue("itaxrate", "17")   '税率，Double类型
    '        domHead(0).SetValue("cexch_name", "人民币")   '币种，String类型
    '        domHead(0).SetValue("cmaker", "demo1")   '制单人，String类型
    '        '’ domHead(0).SetValue("breturnflag", "字段值")   '退货标志，String类型
    '        '   domHead(0).SetValue("ufts", "字段值")   '时间戳，String类型
    '        domHead(0).SetValue("cstcode", "01")   '销售类型编号，String类型
    '        domHead(0).SetValue("cdepcode", "07")   '部门编码，String类型
    '        domHead(0).SetValue("ccuscode", "01")   '客户编码，String类型
    '        'domHead(0).SetValue("ccushand", "字段值")   '客户联系人手机，String类型
    '        'domHead(0).SetValue("cpsnophone", "字段值")   '业务员办公电话，String类型
    '        'domHead(0).SetValue("cpsnmobilephone", "字段值")   '业务员手机，String类型
    '        'domHead(0).SetValue("cattachment", "字段值")   '附件，String类型
    '        'domHead(0).SetValue("csscode", "字段值")   '结算方式编码，String类型
    '        'domHead(0).SetValue("cssname", "字段值")   '结算方式，String类型
    '        'domHead(0).SetValue("cinvoicecompany", "字段值")   '开票单位编码，String类型
    '        'domHead(0).SetValue("cinvoicecompanyabbname", "字段值")   '开票单位简称，String类型
    '        'domHead(0).SetValue("ccuspersoncode", "字段值")   '联系人编码，String类型
    '        'domHead(0).SetValue("dclosedate", "字段值")   '关闭日期，String类型
    '        'domHead(0).SetValue("dclosesystime", "字段值")   '关闭时间，String类型
    '        'domHead(0).SetValue("bmustbook", "字段值")   '必有定金，String类型
    '        'domHead(0).SetValue("fbookratio", "字段值")   '定金比例，String类型
    '        'domHead(0).SetValue("cgathingcode", "字段值")   '收款单号，String类型
    '        'domHead(0).SetValue("fbooksum", "字段值")   '定金原币金额，String类型
    '        'domHead(0).SetValue("fbooknatsum", "字段值")   '定金本币金额，String类型
    '        'domHead(0).SetValue("fgbooknatsum", "字段值")   '定金累计实收本币金额，String类型
    '        'domHead(0).SetValue("fgbooksum", "字段值")   '定金累计实收原币金额，String类型
    '        'domHead(0).SetValue("ccrmpersonname", "字段值")   '相关员工，String类型
    '        'domHead(0).SetValue("csysbarcode", "字段值")   '单据条码，String类型
    '        'domHead(0).SetValue("ioppid", "字段值")   '销售机会ID，String类型
    '        'domHead(0).SetValue("contract_status", "字段值")   'contract_status，String类型
    '        'domHead(0).SetValue("csvouchtype", "字段值")   '来源电商，String类型
    '        'domHead(0).SetValue("bcashsale", "字段值")   '现款结算，String类型
    '        'domHead(0).SetValue("iflowid", "字段值")   '流程id，String类型
    '        'domHead(0).SetValue("cflowname", "字段值")   '流程分支描述，String类型
    '        'domHead(0).SetValue("cchangeverifier", "字段值")   '变更审批人，String类型
    '        'domHead(0).SetValue("dchangeverifydate", "字段值")   '变更审批日期，String类型
    '        'domHead(0).SetValue("dchangeverifytime", "字段值")   '变更审批时间，String类型

    '        '***************************** 以下是非必输字段 ****************************
    '        'domHead(0).SetValue("fstockquanO", "字段值")   '现存件数，Double类型
    '        'domHead(0).SetValue("fcanusequanO", "字段值")   '可用件数，Double类型
    '        'domHead(0).SetValue("iverifystate", "字段值")   'iverifystate，String类型
    '        'domHead(0).SetValue("ireturncount", "字段值")   'ireturncount，String类型
    '        'domHead(0).SetValue("icreditstate", "字段值")   'icreditstate，String类型
    '        'domHead(0).SetValue("iswfcontrolled", "字段值")   'iswfcontrolled，String类型
    '        'domHead(0).SetValue("dpredatebt", "字段值")   '预发货日期，Date类型
    '        'domHead(0).SetValue("dpremodatebt", "字段值")   '预完工日期，Date类型
    '        'domHead(0).SetValue("caddcode", "字段值")   '收货地址编码，String类型
    '        'domHead(0).SetValue("cdeliverunit", "字段值")   '收货单位，String类型
    '        'domHead(0).SetValue("ccontactname", "字段值")   '收货联系人，String类型
    '        'domHead(0).SetValue("cofficephone", "字段值")   '收货联系电话，String类型
    '        'domHead(0).SetValue("cmobilephone", "字段值")   '收货联系人手机，String类型
    '        'domHead(0).SetValue("cpayname", "字段值")   '付款条件，String类型
    '        'domHead(0).SetValue("cpersonname", "字段值")   '业 务 员，String类型
    '        domHead(0).SetValue("iexchrate", "1")   '汇率，Double类型
    '        'domHead(0).SetValue("cmemo", "字段值")   '备    注，String类型
    '        'domHead(0).SetValue("cverifier", "字段值")   '审核人，String类型
    '        'domHead(0).SetValue("ccloser", "字段值")   '关闭人，String类型
    '        'domHead(0).SetValue("clocker", "字段值")   '锁定人，String类型
    '        'domHead(0).SetValue("ivtid", "字段值")   '单据模版号，Integer类型
    '        'domHead(0).SetValue("istatus", "字段值")   '订单状态，String类型
    '        'domHead(0).SetValue("ccrechppass", "字段值")   '信用审核口令，String类型
    '        'domHead(0).SetValue("clowpricepass", "字段值")   '最低售价口令，String类型
    '        'domHead(0).SetValue("bcontinue", "字段值")   '是否继续，String类型
    '        'domHead(0).SetValue("isumx", "字段值")   '价税合计，Double类型
    '        'domHead(0).SetValue("zdsum", "字段值")   '整单合计，Double类型
    '        'domHead(0).SetValue("ccusname", "字段值")   '客户名称，String类型
    '        'domHead(0).SetValue("ccusphone", "字段值")   '联系电话，String类型
    '        'domHead(0).SetValue("csccode", "字段值")   '发运方式编码，String类型
    '        'domHead(0).SetValue("cpaycode", "字段值")   '付款条件编码，String类型
    '        'domHead(0).SetValue("ccusperson", "字段值")   '联系人，String类型
    '        'domHead(0).SetValue("coppcode", "字段值")   '商机编码，String类型
    '        'domHead(0).SetValue("ccusaddress", "字段值")   '客户地址，String类型
    '        'domHead(0).SetValue("cpersoncode", "字段值")   '业务员编码，String类型
    '        'domHead(0).SetValue("iarmoney", "字段值")   '客户应收余额，Double类型
    '        'domHead(0).SetValue("ccusoaddress", "字段值")   '发货地址，String类型
    '        'domHead(0).SetValue("imoney", "字段值")   '定    金，Double类型
    '        'domHead(0).SetValue("cscname", "字段值")   '发运方式，String类型
    '        'domHead(0).SetValue("cchanger", "字段值")   '变更人，String类型
    '        'domHead(0).SetValue("dcreatesystime", "字段值")   '制单时间，Date类型
    '        'domHead(0).SetValue("dverifysystime", "字段值")   '审核时间，Date类型
    '        'domHead(0).SetValue("dmodifysystime", "字段值")   '修改时间，Date类型
    '        'domHead(0).SetValue("cmodifier", "字段值")   '修改人，String类型
    '        'domHead(0).SetValue("dmoddate", "字段值")   '修改日期，Date类型
    '        'domHead(0).SetValue("dverifydate", "字段值")   '审核日期，Date类型
    '        'domHead(0).SetValue("cdefine16", "字段值")   '表头自定义项16，Double类型
    '        'domHead(0).SetValue("ccrechpname", "字段值")   '信用审核人，String类型
    '        'domHead(0).SetValue("ccusdefine1", "字段值")   '客户自定义项1，String类型
    '        'domHead(0).SetValue("ccusdefine2", "字段值")   '客户自定义项2，String类型
    '        'domHead(0).SetValue("ccusdefine3", "字段值")   '客户自定义项3，String类型
    '        'domHead(0).SetValue("ccusdefine4", "字段值")   '客户自定义项4，String类型
    '        'domHead(0).SetValue("zdsumdx", "字段值")   '整单合计（大写），String类型
    '        'domHead(0).SetValue("isumdx", "字段值")   '价税合计（大写），String类型
    '        'domHead(0).SetValue("ccusdefine5", "字段值")   '客户自定义项5，String类型
    '        'domHead(0).SetValue("ccusdefine6", "字段值")   '客户自定义项6，String类型
    '        'domHead(0).SetValue("ccusdefine7", "字段值")   '客户自定义项7，String类型
    '        'domHead(0).SetValue("ccusdefine8", "字段值")   '客户自定义项8，String类型
    '        'domHead(0).SetValue("ccusdefine9", "字段值")   '客户自定义项9，String类型
    '        'domHead(0).SetValue("ccusdefine10", "字段值")   '客户自定义项10，String类型
    '        'domHead(0).SetValue("ccusdefine11", "字段值")   '客户自定义项11，String类型
    '        'domHead(0).SetValue("ccusdefine12", "字段值")   '客户自定义项12，String类型
    '        'domHead(0).SetValue("ccusdefine13", "字段值")   '客户自定义项13，String类型
    '        'domHead(0).SetValue("ccusdefine14", "字段值")   '客户自定义项14，String类型
    '        'domHead(0).SetValue("ccusdefine15", "字段值")   '客户自定义项15，String类型
    '        'domHead(0).SetValue("ccusdefine16", "字段值")   '客户自定义项16，String类型
    '        'domHead(0).SetValue("icuscreline", "字段值")   '用户信用度，Double类型
    '        'domHead(0).SetValue("fstockquan", "字段值")   '现存数量，Double类型
    '        'domHead(0).SetValue("fcanusequan", "字段值")   '可用数量，Double类型
    '        'domHead(0).SetValue("cdefine1", "字段值")   '表头自定义项1，String类型
    '        'domHead(0).SetValue("cdefine2", "字段值")   '表头自定义项2，String类型
    '        'domHead(0).SetValue("cdefine3", "字段值")   '表头自定义项3，String类型
    '        'domHead(0).SetValue("cdefine4", "字段值")   '表头自定义项4，Date类型
    '        'domHead(0).SetValue("cdefine5", "字段值")   '表头自定义项5，Integer类型
    '        'domHead(0).SetValue("cdefine6", "字段值")   '表头自定义项6，Date类型
    '        'domHead(0).SetValue("cdefine7", "字段值")   '表头自定义项7，Double类型
    '        'domHead(0).SetValue("cdefine8", "字段值")   '表头自定义项8，String类型
    '        'domHead(0).SetValue("cdefine9", "字段值")   '表头自定义项9，String类型
    '        'domHead(0).SetValue("cdefine10", "字段值")   '表头自定义项10，String类型
    '        'domHead(0).SetValue("cdefine11", "字段值")   '表头自定义项11，String类型
    '        'domHead(0).SetValue("cdefine12", "字段值")   '表头自定义项12，String类型
    '        'domHead(0).SetValue("cdefine13", "字段值")   '表头自定义项13，String类型
    '        'domHead(0).SetValue("cdefine14", "字段值")   '表头自定义项14，String类型
    '        'domHead(0).SetValue("cdefine15", "字段值")   '表头自定义项15，Integer类型
    '        'domHead(0).SetValue("ccreditcuscode", "字段值")   '信用单位编码，String类型
    '        'domHead(0).SetValue("ccreditcusname", "字段值")   '信用单位名称，String类型
    '        'domHead(0).SetValue("cgatheringplan", "字段值")   '收付款协议编码，String类型
    '        'domHead(0).SetValue("cgatheringplanname", "字段值")   '收付款协议名称，String类型

    '        '给BO表体参数domBody赋值，此BO参数的业务类型为销售订单，属表体参数。BO参数均按引用传递
    '        '提示：给BO表体参数domBody赋值有两种方法

    '        '方法一是直接传入MSXML2.DOMDocument对象
    '        'Dim domBody As New MSXML2.DOMDocument
    '        'u8apiBroker.AssignNormalValue "domBody", domBody

    '        '方法二是构造BusinessObject对象，具体方法如下：
    '        Dim domBody As BusinessObject
    '        domBody = u8apiBroker.GetBoParam("domBody")
    '        domBody.RowCount = 10 '设置BO对象(表体)行数为多行
    '        '可以自由设置BO对象(表体)行数为任意大于零的整数，也可以不设置而自动增加
    '        '给BO对象(表体)的字段赋值，值可以是真实类型，也可以是无类型字符串
    '        '以下代码示例只设置第一行值。各字段定义详见API服务接口定义

    '        '****************************** 以下是必输字段 *****************************
    '        '’  domBody(0).SetValue("isosid", "字段值")   '主关键字段，Integer类型
    '        domBody(0).SetValue("cinvname", "排尘钢管总成")   '存货名称，String类型
    '        domBody(0).SetValue("cinvcode", "1030000001")   '存货编码，String类型
    '        '  domBody(0).SetValue("autoid", "字段值")   '销售订单 2，Integer类型
    '        domBody(0).SetValue("iquantity", "19")   '数量，Double类型
    '        domBody(0).SetValue("dpredate", "2016-4-20")   '预发货日期，Date类型
    '        domBody(0).SetValue("dpremodate", "2016-4-19")   '预完工日期，Date类型
    '        'domBody(0).SetValue("borderbom", "字段值")   '是否订单BOM，Integer类型
    '        'domBody(0).SetValue("borderbomover", "字段值")   '订单BOM是否完成，Integer类型
    '        domBody(0).SetValue("id", "100000002")   '主表id，Integer类型
    '        'domBody(0).SetValue("iinvexchrate", "字段值")   '换算率，Double类型
    '        'domBody(0).SetValue("cunitid", "字段值")   '销售单位编码，String类型
    '        'domBody(0).SetValue("cinva_unit", "字段值")   '销售单位，String类型
    '        domBody(0).SetValue("cinvm_unit", "101")   '主计量单位，String类型
    '        'domBody(0).SetValue("igrouptype", "字段值")   '单位类型，Integer类型
    '        domBody(0).SetValue("cgroupcode", "01")   '计量单位组，String类型
    '        'domBody(0).SetValue("dreleasedate", "字段值")   '预留失效日期，Date类型
    '        domBody(0).SetValue("editprop", "A")   '编辑属性：A表新增，M表修改，D表删除，String类型
    '        'domBody(0).SetValue("fstockquano", "字段值")   '现存件数，String类型
    '        'domBody(0).SetValue("fcanusequano", "字段值")   '可用件数，String类型
    '        'domBody(0).SetValue("iimid", "字段值")   '进口订单明细行，String类型
    '        'domBody(0).SetValue("btracksalebill", "字段值")   'PE跟单，String类型
    '        'domBody(0).SetValue("ccorvouchtype", "字段值")   '来源单据类型，String类型
    '        'domBody(0).SetValue("ccorvouchtypename", "字段值")   '来源单据名称，String类型
    '        'domBody(0).SetValue("icorrowno", "字段值")   '来源单据行号，String类型
    '        'domBody(0).SetValue("fcanusequan", "字段值")   '可用量，String类型
    '        'domBody(0).SetValue("fstockquan", "字段值")   '现存量，String类型
    '        'domBody(0).SetValue("bsaleprice", "字段值")   '报价含税，String类型
    '        'domBody(0).SetValue("bgift", "字段值")   '赠品，String类型
    '        'domBody(0).SetValue("forecastdid", "字段值")   '预测单子表ID，String类型
    '        'domBody(0).SetValue("cdetailsdemandcode", "字段值")   '子件需求分类代号，String类型
    '        'domBody(0).SetValue("cdetailsdemandmemo", "字段值")   '子件需求分类说明，String类型
    '        'domBody(0).SetValue("cbsysbarcode", "字段值")   '单据行条码，String类型
    '        'domBody(0).SetValue("busecusbom", "字段值")   '使用客户BOM，String类型
    '        'domBody(0).SetValue("bptomodel", "字段值")   'bptomodel，String类型
    '        'domBody(0).SetValue("cparentcode", "字段值")   '父节点编码，String类型
    '        'domBody(0).SetValue("cchildcode", "字段值")   '子节点编码，String类型
    '        'domBody(0).SetValue("icalctype", "字段值")   '发货模式，String类型
    '        'domBody(0).SetValue("fchildqty", "字段值")   '使用数量，String类型
    '        'domBody(0).SetValue("fchildrate", "字段值")   '权重比例，String类型

    '        ''***************************** 以下是非必输字段 ****************************
    '        'domBody(0).SetValue("natoseqid", "字段值")   'ato行id，String类型
    '        'domBody(0).SetValue("natostatus", "字段值")   'ato行编辑状态，String类型
    '        'domBody(0).SetValue("iquoid", "字段值")   '报价id，String类型
    '        'domBody(0).SetValue("cscloser", "字段值")   '行关闭人，String类型
    '        'domBody(0).SetValue("irowno", "字段值")   '行号，String类型
    '        'domBody(0).SetValue("cconfigstatus", "字段值")   '选配标志，Integer类型
    '        'domBody(0).SetValue("ippartseqid", "字段值")   '选配序号，String类型
    '        'domBody(0).SetValue("cquocode", "字段值")   '报价单号，String类型
    '        'domBody(0).SetValue("cinvstd", "字段值")   '规格型号，String类型
    '        'domBody(0).SetValue("ccontractid", "字段值")   '合同编码，String类型
    '        'domBody(0).SetValue("ccontractrowguid", "字段值")   '合同标的RowGuid，String类型
    '        'domBody(0).SetValue("ccontracttagcode", "字段值")   '合同标的编码，String类型
    '        'domBody(0).SetValue("icusbomid", "字段值")   '客户BomID，String类型
    '        'domBody(0).SetValue("ippartqty", "字段值")   '母件数量，String类型
    '        'domBody(0).SetValue("ippartid", "字段值")   '母件物料ID，String类型
    '        'domBody(0).SetValue("imoquantity", "字段值")   '下达生产量，Double类型
    '        'domBody(0).SetValue("batomodel", "字段值")   '是否ATO件，Integer类型
    '        'domBody(0).SetValue("inum", "字段值")   '件数，Double类型
    '        'domBody(0).SetValue("fsalecost", "字段值")   '零售单价，Double类型
    '        'domBody(0).SetValue("itaxunitprice", "字段值")   '含税单价，Double类型
    '        'domBody(0).SetValue("iquotedprice", "字段值")   '报价，Double类型
    '        'domBody(0).SetValue("iunitprice", "字段值")   '无税单价，Double类型
    '        'domBody(0).SetValue("imoney", "字段值")   '无税金额，Double类型
    '        'domBody(0).SetValue("itax", "字段值")   '税额，Double类型
    '        'domBody(0).SetValue("isum", "字段值")   '价税合计，Double类型
    '        'domBody(0).SetValue("fsaleprice", "字段值")   '零售金额，Double类型
    '        'domBody(0).SetValue("inatunitprice", "字段值")   '本币单价，Double类型
    '        'domBody(0).SetValue("inatmoney", "字段值")   '本币金额，Double类型
    '        'domBody(0).SetValue("inattax", "字段值")   '本币税额，Double类型
    '        'domBody(0).SetValue("inatsum", "字段值")   '本币价税合计，Double类型
    '        'domBody(0).SetValue("inatdiscount", "字段值")   '本币折扣额，Double类型
    '        'domBody(0).SetValue("idiscount", "字段值")   '折扣额，Double类型
    '        'domBody(0).SetValue("ifhquantity", "字段值")   '发货数量，Double类型
    '        'domBody(0).SetValue("ifhnum", "字段值")   '发货件数，Double类型
    '        'domBody(0).SetValue("ifhmoney", "字段值")   '发货金额，Double类型
    '        'domBody(0).SetValue("ikpquantity", "字段值")   '开票数量，Double类型
    '        'domBody(0).SetValue("ikpnum", "字段值")   '开票件数，Double类型
    '        'domBody(0).SetValue("ikpmoney", "字段值")   '开票金额，Double类型
    '        'domBody(0).SetValue("iinvlscost", "字段值")   '最低售价，Double类型
    '        'domBody(0).SetValue("cfree1", "字段值")   '自由项1，String类型
    '        'domBody(0).SetValue("cfree2", "字段值")   '自由项2，String类型
    '        'domBody(0).SetValue("bservice", "字段值")   '是否应税劳务，String类型
    '        'domBody(0).SetValue("bfree1", "字段值")   '是否有自由项1，String类型
    '        'domBody(0).SetValue("bfree2", "字段值")   '是否有自由项2，String类型
    '        'domBody(0).SetValue("bfree3", "字段值")   '是否有自由项3，String类型
    '        'domBody(0).SetValue("bfree4", "字段值")   '是否有自由项4，String类型
    '        'domBody(0).SetValue("bfree5", "字段值")   '是否有自由项5，String类型
    '        'domBody(0).SetValue("bfree6", "字段值")   '是否有自由项6，String类型
    '        'domBody(0).SetValue("bfree7", "字段值")   '是否有自由项7，String类型
    '        'domBody(0).SetValue("bfree8", "字段值")   '是否有自由项8，String类型
    '        'domBody(0).SetValue("bfree9", "字段值")   '是否有自由项9，String类型
    '        'domBody(0).SetValue("bfree10", "字段值")   '是否有自由项10，String类型
    '        'domBody(0).SetValue("cmemo", "字段值")   '备注，String类型
    '        'domBody(0).SetValue("cinvdefine1", "字段值")   '存货自定义项1，String类型
    '        'domBody(0).SetValue("cinvdefine4", "字段值")   '存货自定义项4，String类型
    '        'domBody(0).SetValue("cinvdefine5", "字段值")   '存货自定义项5，String类型
    '        'domBody(0).SetValue("cinvdefine6", "字段值")   '存货自定义项6，String类型
    '        'domBody(0).SetValue("cinvdefine7", "字段值")   '存货自定义项7，String类型
    '        'domBody(0).SetValue("bsalepricefree1", "字段值")   '是否自由项定价1，String类型
    '        'domBody(0).SetValue("bsalepricefree2", "字段值")   '是否自由项定价2，String类型
    '        'domBody(0).SetValue("bsalepricefree3", "字段值")   '是否自由项定价3，String类型
    '        'domBody(0).SetValue("bsalepricefree4", "字段值")   '是否自由项定价4，String类型
    '        'domBody(0).SetValue("bsalepricefree5", "字段值")   '是否自由项定价5，String类型
    '        'domBody(0).SetValue("bsalepricefree6", "字段值")   '是否自由项定价6，String类型
    '        'domBody(0).SetValue("bsalepricefree7", "字段值")   '是否自由项定价7，String类型
    '        'domBody(0).SetValue("bsalepricefree8", "字段值")   '是否自由项定价8，String类型
    '        'domBody(0).SetValue("bsalepricefree9", "字段值")   '是否自由项定价9，String类型
    '        'domBody(0).SetValue("bsalepricefree10", "字段值")   '是否自由项定价10，String类型
    '        'domBody(0).SetValue("iaoids", "字段值")   '預訂單子表id，Integer类型
    '        'domBody(0).SetValue("cpreordercode", "字段值")   '预订单号，Integer类型
    '        'domBody(0).SetValue("idemandtype", "字段值")   '需求跟踪方式，Integer类型
    '        'domBody(0).SetValue("cdemandcode", "字段值")   '需求分类代号，String类型
    '        'domBody(0).SetValue("cdemandmemo", "字段值")   '需求分类说明，String类型
    '        'domBody(0).SetValue("cinvdefine8", "字段值")   '存货自定义项8，String类型
    '        'domBody(0).SetValue("cinvdefine9", "字段值")   '存货自定义项9，String类型
    '        'domBody(0).SetValue("cinvdefine10", "字段值")   '存货自定义项10，String类型
    '        'domBody(0).SetValue("cinvdefine11", "字段值")   '存货自定义项11，String类型
    '        'domBody(0).SetValue("cinvdefine12", "字段值")   '存货自定义项12，String类型
    '        'domBody(0).SetValue("cinvdefine13", "字段值")   '存货自定义项13，String类型
    '        'domBody(0).SetValue("cinvdefine14", "字段值")   '存货自定义项14，String类型
    '        'domBody(0).SetValue("cinvdefine15", "字段值")   '存货自定义项15，String类型
    '        'domBody(0).SetValue("cinvdefine16", "字段值")   '存货自定义项16，String类型
    '        'domBody(0).SetValue("cinvdefine2", "字段值")   '存货自定义项2，String类型
    '        'domBody(0).SetValue("cinvdefine3", "字段值")   '存货自定义项3，String类型
    '        'domBody(0).SetValue("binvtype", "字段值")   '存货类型，String类型
    '        'domBody(0).SetValue("cdefine22", "字段值")   '表体自定义项1，String类型
    '        'domBody(0).SetValue("cdefine23", "字段值")   '表体自定义项2，String类型
    '        'domBody(0).SetValue("cdefine24", "字段值")   '表体自定义项3，String类型
    '        'domBody(0).SetValue("cdefine25", "字段值")   '表体自定义项4，String类型
    '        'domBody(0).SetValue("cdefine26", "字段值")   '表体自定义项5，Double类型
    '        'domBody(0).SetValue("cdefine27", "字段值")   '表体自定义项6，Double类型
    '        'domBody(0).SetValue("itaxrate", "字段值")   '税率（％），Double类型
    '        'domBody(0).SetValue("kl2", "字段值")   '扣率2（％），Double类型
    '        'domBody(0).SetValue("citemcode", "字段值")   '项目编码，String类型
    '        'domBody(0).SetValue("citem_class", "字段值")   '项目大类编码，String类型
    '        'domBody(0).SetValue("dkl1", "字段值")   '倒扣1（％），Double类型
    '        'domBody(0).SetValue("dkl2", "字段值")   '倒扣2（％），Double类型
    '        'domBody(0).SetValue("citemname", "字段值")   '项目名称，String类型
    '        'domBody(0).SetValue("citem_cname", "字段值")   '项目大类名称，String类型
    '        'domBody(0).SetValue("cfree3", "字段值")   '自由项3，String类型
    '        'domBody(0).SetValue("cfree4", "字段值")   '自由项4，String类型
    '        'domBody(0).SetValue("cfree5", "字段值")   '自由项5，String类型
    '        'domBody(0).SetValue("cfree6", "字段值")   '自由项6，String类型
    '        'domBody(0).SetValue("cfree7", "字段值")   '自由项7，String类型
    '        'domBody(0).SetValue("cfree8", "字段值")   '自由项8，String类型
    '        'domBody(0).SetValue("cfree9", "字段值")   '自由项9，String类型
    '        'domBody(0).SetValue("cfree10", "字段值")   '自由项10，String类型
    '        'domBody(0).SetValue("cdefine28", "字段值")   '表体自定义项7，String类型
    '        'domBody(0).SetValue("cdefine29", "字段值")   '表体自定义项8，String类型
    '        'domBody(0).SetValue("cdefine30", "字段值")   '表体自定义项9，String类型
    '        'domBody(0).SetValue("cdefine31", "字段值")   '表体自定义项10，String类型
    '        'domBody(0).SetValue("cdefine32", "字段值")   '表体自定义项11，String类型
    '        'domBody(0).SetValue("corufts", "字段值")   '对应单据时间戳，String类型
    '        'domBody(0).SetValue("cdefine33", "字段值")   '表体自定义项12，String类型
    '        'domBody(0).SetValue("cdefine34", "字段值")   '表体自定义项13，Integer类型
    '        'domBody(0).SetValue("cdefine35", "字段值")   '表体自定义项14，Integer类型
    '        'domBody(0).SetValue("cdefine36", "字段值")   '表体自定义项15，Date类型
    '        'domBody(0).SetValue("cdefine37", "字段值")   '表体自定义项16，Date类型
    '        'domBody(0).SetValue("binvmodel", "字段值")   '是否模型件，Integer类型
    '        'domBody(0).SetValue("csrpolicy", "字段值")   '供需政策，String类型
    '        'domBody(0).SetValue("iprekeepquantity", "字段值")   '预留数量，Double类型
    '        'domBody(0).SetValue("iprekeepnum", "字段值")   '预留件数，Double类型
    '        'domBody(0).SetValue("iprekeeptotquantity", "字段值")   '预留母件和子件数量，Double类型
    '        'domBody(0).SetValue("iprekeeptotnum", "字段值")   '预留母件子件件数，Double类型
    '        'domBody(0).SetValue("fcusminprice", "字段值")   '客户最低售价，Double类型
    '        'domBody(0).SetValue("ccusinvcode", "字段值")   '客户存货编码，String类型
    '        'domBody(0).SetValue("ccusinvname", "字段值")   '客户存货名称，String类型
    '        'domBody(0).SetValue("cinvaddcode", "字段值")   '存货代码，String类型
    '        'domBody(0).SetValue("dbclosedate", "字段值")   '关闭日期，Date类型
    '        'domBody(0).SetValue("dbclosesystime", "字段值")   '关闭时间，Date类型
    '        'domBody(0).SetValue("kl", "字段值")   '扣率（％），Double类型

    '        '给普通参数VoucherState赋值。此参数的数据类型为Integer，此参数按值传递，表示状态:0增加;1修改
    '        u8apiBroker.AssignNormalValue("VoucherState", 0)  '参数类型：Integer

    '        '该参数vNewID为INOUT型普通参数。此参数的数据类型为String，此参数按值传递。在API调用返回时，可以通过GetResult("vNewID")获取其值
    '        u8apiBroker.AssignNormalValue("vNewID", "000000002")  '参数类型：String

    '        '给普通参数DomConfig赋值。此参数的数据类型为MSXML2.IXMLDOMDocument2，此参数按引用传递，表示ATO,PTO选配
    '        Dim CurDom As New DOMDocument
    '        u8apiBroker.AssignNormalValue("DomConfig", CurDom)  '参数类型：MSXML2.IXMLDOMDocument2

    '        '第五步：调用API
    '        If u8apiBroker.InvokeApi() = False Then

    '            '第六步：错误处理
    '            If u8apiBroker.ErrorType = ExceptionType.Business Then

    '                '处理API业务错误
    '            ElseIf u8apiBroker.ErrorType = ExceptionType.System Then

    '                '处理系统错误
    '            End If
    '        Else
    '            '第七步：获取返回结果

    '            '获取返回值
    '            '获取普通返回值。此返回值数据类型为String，此参数按值传递，表示成功为空串
    '            Dim result As String
    '            result = CStr(u8apiBroker.GetReturnValue())
    '            '获取out/inout参数值

    '            '获取普通INOUT参数vNewID。此返回值数据类型为String，在使用该参数之前，请判断是否为空
    '            Dim vNewIDRet As String
    '            vNewIDRet = CStr(u8apiBroker.GetResult("vNewID"))

    '        End If

    '        '结束本次调用，释放API资源
    '        u8apiBroker.Disconnect()

    '        u8apiBroker = Nothing
    '        MsgBox("OK")
    '        Exit Sub
    'ErrHandler:
    '        MsgBox(Err.Description)


    '    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim x As String
        If Not Is64bit() Then
            x = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Excel", "TypeGuessRows", Nothing)
            If x <> "0" Then
                My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Excel", "TypeGuessRows", 0, Microsoft.Win32.RegistryValueKind.DWord)
            End If
        ElseIf Is64bit() Then
            x = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Excel", "TypeGuessRows", Nothing)
            If x <> "0" Then
                My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Excel", "TypeGuessRows", 0, Microsoft.Win32.RegistryValueKind.DWord)
            End If
        End If

        Dim dt As DataTable


        Try
            Dim _Connectstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<FilePath>;Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
            excConn = New OleDb.OleDbConnection(_Connectstring.Replace("<FilePath>", filename))
            excConn.Open()
            '   MsgBox(_Connectstring)
            Dim mydataset As DataSet = New DataSet
            Using da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("select * from [Sheet1$] ", excConn)

                Try
                    dt = New DataTable
                    da.Fill(dt)
                Catch ex As Exception
                    '   Console.WriteLine(ex)
                    MsgBox("请注意Sheet名是否为Sheet1！")
                End Try

                '   dt = mydataset.Tables("Sheet1")

                DataGridView1.AutoGenerateColumns = True '自动创建列  
                DataGridView1.DataSource = dt
            End Using

        Catch ex As Exception

        End Try
     
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim x As String = TextBox1.Text
        Dim y
        y = CDate(x)
        MsgBox(y)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim fd(6) As String
        filename = "C:/123.xls"
        Dim _Connectstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<FilePath>;Extended Properties=""Excel 8.0;HDR=NO;IMEX=1"""
        excConn = New OleDb.OleDbConnection(_Connectstring.Replace("<FilePath>", filename))
        excConn.Open()
        Dim dt As DataTable = New DataTable()
        Dim Sql As String = "select * from [Sheet1$] where F1 is not null"
        Dim mycommand As OleDbDataAdapter = New OleDbDataAdapter(Sql, excConn)
        mycommand.Fill(dt)
        fd(0) = dt.Rows(0)("F3").ToString
        fd(1) = dt.Rows(0)("F4").ToString
        fd(2) = dt.Rows(0)("F5").ToString
        fd(3) = dt.Rows(0)("F6").ToString
        fd(4) = dt.Rows(0)("F7").ToString
        fd(5) = dt.Rows(0)("F8").ToString

        Dim dv As DataView = New DataView(dt)
        Dim dt2 As DataTable = dv.ToTable(True, "F1")

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim x As Date = TextBox1.Text
        Dim y As Integer = 0 - CInt(TextBox3.Text)
        x = DateAdd("d", y, x)
        TextBox2.Text = x
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        On Error GoTo ErrHandler

        ''第一步：构造u8login对象并登陆
        ''如果当前环境中有login对象则可以省去第一步
        'Dim ologin As Object
        'Set ologin = CreateObject("U8Login.clsLogin")
        'If Not ologin.login("AS", strAccID, strYear, strUserID, strPwd, strDate, strServer) Then

        '第二步：构造环境上下文对象，传入login，并按需设置其它上下文参数
        Dim u8EnvCtx As New U8EnvContext
        u8EnvCtx.U8Login = u8login

        '采购所有接口均支持内部独立事务和外部事务，默认内部事务
        '如果是外部事务，则需要传递ADO.Connection对象，并将IsIndependenceTransaction设置为false
        'Dim bizDbConn As New ADO.Connection
        'Set u8EnvCtx.BizDbConnection = bizDbConn
        'u8EnvCtx.IsIndependenceTransaction = false

        '设置上下文参数
        u8EnvCtx.SetApiContext("VoucherType", "1")       '上下文数据类型：int，含义：单据类型，采购订单 1
        u8EnvCtx.SetApiContext("bPositive", "True")      '上下文数据类型：bool，含义：红蓝标识：True,蓝字
        u8EnvCtx.SetApiContext("sBillType", "")          '上下文数据类型：string，含义：为空串
        u8EnvCtx.SetApiContext("sBusType", "普通采购")   '上下文数据类型：string，含义：业务类型：普通采购,直运采购,受托代销

        '第三步：构造ApiBroker对象,调用Connect,传入Api的地址标识(Url)，传入上下文
        Dim u8apiBroker As New U8ApiComBroker
        u8apiBroker.Connect("U8API/PurchaseOrder/VoucherSave", u8EnvCtx)

        '第四步：API参数赋值

        '给BO表头参数DomHead赋值，此BO参数的业务类型为采购订单，属表头参数。BO参数均按引用传递
        '提示：给BO表头参数DomHead赋值有两种方法
        '方法一是直接传入MSXML2.DOMDocument对象
        Dim domHead As New MSXML2.DOMDocument   '单据表头DOM
        Dim domBody As New MSXML2.DOMDocument   '单据表体DOM
        Dim eleHead As IXMLDOMElement
        Dim eleBody As IXMLDOMElement
        Dim ele As IXMLDOMElement
        Dim rs As New ADODB.Recordset
        Dim strSQL As String
        Dim strPOID As String
        Dim strCode As String
        'Dim strSysBarCode As String
        'Dim strBSysBarCode As String
        Dim i As Integer
        Dim id As Long

        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        '查询采购订单表头视图zpurpoheader，获取表头DOM结构
        '如果有表头扩展自定义项，则可以关联PO_Pomain_extradefine表
        'editprop（单据编辑属性）：A表新增单据，M表修改单据，D表删除单据
        '新增时只需要一个空的DOM结构，所以查询条件为where 1=0
        strSQL = "select cast(null as nvarchar(2)) as editprop,* from zpurpoheader where 1=0"
        rs.Open(strSQL, conn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        rs.Save(domHead, ADODB.PersistFormatEnum.adPersistXML)
        rs.Close()

        '增加表头数据节点z:row
        eleHead = domHead.selectSingleNode("//rs:data")
        ele = domHead.createElement("z:row")
        eleHead.appendChild(ele)

        strPOID = "0"
        strCode = TextBox1.Text
        'strSysBarCode = "||pupo|" & strCode

        '给表头DOM赋值
        setAttribute(ele, "editprop", "A")               '编辑属性：A表新增，M表修改，D表删除，String类型
        setAttribute(ele, "ivtid", "8173")               '单据模版号，Integer类型
        setAttribute(ele, "poid", strPOID)               '主关键字段，Integer类型
        setAttribute(ele, "cbustype", "普通采购")        '业务类型，Integer类型
        setAttribute(ele, "dpodate", "2016-06-14")       '日期，Date类型
        setAttribute(ele, "cpoid", strCode)              '订单编号，String类型
        setAttribute(ele, "cvencode", "01")               '供货单位编号，String类型
        setAttribute(ele, "cvenname", "十堰鼎汇贸易有限公司")              '供应商全称，String类型
        setAttribute(ele, "cvenabbname", "鼎汇")            '供货单位，String类型

        setAttribute(ele, "cexch_name", "人民币")        '币种，String类型
        setAttribute(ele, "nflat", "1")                  '汇率，Double类型
        setAttribute(ele, "itaxrate", "17")              '税率，Double类型
        setAttribute(ele, "idiscounttaxtype", "0")       '扣税类别，Integer类型

        'setAttribute ele, "csysbarcode", strSysBarCode  '单据条码，String类型
        setAttribute(ele, "cmaker", "demo")              '制单人，String类型

        setAttribute(ele, "chdefine1", "abc")            '表头扩展自定义项1，String类型

        u8apiBroker.AssignNormalValue("DomHead", domHead)

        '方法二是构造BusinessObject对象，具体方法如下：
        'Dim DomHead As BusinessObject
        'Set DomHead = u8apiBroker.GetBoParam("DomHead")
        'DomHead.RowCount = 1 '设置BO对象(表头)行数，只能为一行
        '给BO对象(表头)的字段赋值，值可以是真实类型，也可以是无类型字符串


        '给BO表体参数domBody赋值，此BO参数的业务类型为采购订单，属表体参数。BO参数均按引用传递
        '提示：给BO表体参数domBody赋值有两种方法
        '方法一是直接传入MSXML2.DOMDocument对象
        'Dim domBody As New MSXML2.DOMDocument

        '查询采购订单表体视图zpurpotail，获取表体DOM结构
        '如果有表体扩展自定义项，则可以关联PO_Podetails_extradefine表
        'editprop（单据编辑属性）：A表新增单据，M表修改单据，D表删除单据
        '新增时只需要一个空的DOM结构，所以查询条件为where 1=0
        strSQL = "select cast(null as nvarchar(2)) as editprop,* from zpurpotail  where 1=0"
        rs.Open(strSQL, conn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        rs.Save(domBody, ADODB.PersistFormatEnum.adPersistXML)
        rs.Close()
        rs = Nothing

        id = 1000000036

        '增加表体数据节点z:row
        eleBody = domBody.selectSingleNode("//rs:data")
        '增加两行表体，为了方便编码，两行存货相同，数量不同
        '只是示例，所以不考虑单价和金额字段
        For i = 1 To 2
            ele = domBody.createElement("z:row")
            eleBody.appendChild(ele)

            id = id + i
            'strBSysBarCode = strSysBarCode & "|" & i

            '给表体DOM赋值
            setAttribute(ele, "editprop", "A")                   '编辑属性：A表新增，M表修改，D表删除，String类型
            setAttribute(ele, "id", id)                          '主关键字段，Integer类型
            setAttribute(ele, "poid", strPOID)                   '主表id，Integer类型
            setAttribute(ele, "cinvcode", "3010000003")                 '存货编码，String类型
            setAttribute(ele, "cinvname", "热板")                '存货名称，String类型
            setAttribute(ele, "iquantity", 10 * i)               '数量，Double类型
            setAttribute(ele, "inum", i)                         '件数，Double类型
            setAttribute(ele, "iinvexchrate", "10")              '换算率，Double类型

            setAttribute(ele, "ccomunitcode", "1")               '主计量编码，String类型
            setAttribute(ele, "cinvm_unit", "1")                 '主计量，String类型
            setAttribute(ele, "cunitid", "10")                   '采购单位编码，String类型
            setAttribute(ele, "cinva_unit", "10")                '采购单位，String类型
            setAttribute(ele, "igrouptype", "1")                 '分组类型，String类型

            setAttribute(ele, "darrivedate", "9999-09-30")       '计划到货日期，Date类型
            setAttribute(ele, "ipertaxrate", "17.000000")        '税率，Double类型
            setAttribute(ele, "btaxcost", "0")                   '单价标准，String类型

            setAttribute(ele, "bgsp", "0")                       '是否检验，Integer类型
            setAttribute(ele, "sotype", "0")                     '需求跟踪方式，Integer类型
            setAttribute(ele, "iordertype", "0")                 '销售订单类型，Integer类型

            setAttribute(ele, "ivouchrowno", i)                  '行号，String类型
            'setAttribute ele, "cbsysbarcode", strBSysBarCode    '单据行条码，String类型

            setAttribute(ele, "cbdefine1", "def" & i)            '表体扩展自定义项1，String类型
        Next i


        u8apiBroker.AssignNormalValue("domBody", domBody)

        '方法二是构造BusinessObject对象，具体方法如下：
        'Dim domBody As BusinessObject
        'Set domBody = u8apiBroker.GetBoParam("domBody")
        'domBody.RowCount = 10 '设置BO对象(表体)行数为多行
        '可以自由设置BO对象(表体)行数为任意大于零的整数，也可以不设置而自动增加
        '给BO对象(表体)的字段赋值，值可以是真实类型，也可以是无类型字符串
        '以下代码示例只设置第一行值。各字段定义详见API服务接口定义


        '给普通参数VoucherState赋值。此参数的数据类型为Integer，此参数按值传递，表示单据状态：2新增，1修改，0非编辑
        u8apiBroker.AssignNormalValue("VoucherState", 2)  '参数类型：Integer

        '该参数curID为OUT型参数，由于其数据类型为String，为一般值类型，因此不必传入一个参数变量。在API调用返回时，可以通过GetResult("curID")获取其值

        '该参数CurDom为OUT型参数，由于其数据类型为MSXML2.IXMLDOMDocument2，非一般值类型，因此必须传入一个参数变量。在API调用返回时，可以直接使用该参数
        Dim CurDom As New DOMDocument
        u8apiBroker.AssignNormalValue("CurDom", CurDom)  '参数类型：MSXML2.IXMLDOMDocument2

        '给普通参数UserMode赋值。此参数的数据类型为Integer，此参数按值传递，表示模式，0：CS;1:BS
        u8apiBroker.AssignNormalValue("UserMode", 0)  '参数类型：Integer

        '第五步：调用API
        If u8apiBroker.InvokeApi() = False Then
            '第六步：错误处理
            MsgBox(u8apiBroker.GetLastError())
            If u8apiBroker.ErrorType = ExceptionType.Business Then
                '处理API业务错误
            ElseIf u8apiBroker.ErrorType = ExceptionType.System Then
                '处理系统错误
            End If
        Else
            '第七步：获取返回结果

            '获取返回值
            '获取普通返回值。此返回值数据类型为String，此参数按值传递，表示错误描述：空，正确；非空，错误
            Dim result As String
            result = CStr(u8apiBroker.GetReturnValue())

            '获取out/inout参数值

            '获取普通OUT参数curID。此返回值数据类型为String，在使用该参数之前，请判断是否为空
            Dim curIDRet As String
            curIDRet = CStr(u8apiBroker.GetResult("curID"))

            '获取普通OUT参数CurDom。此返回值数据类型为MSXML2.IXMLDOMDocument2，前面已定义该参数，请直接使用

            If result = "" Then
                MsgBox("新增保存成功！", vbInformation, "提示")
                '   getLastCode()
            Else
                MsgBox(result, vbInformation, "提示")
            End If
        End If

ExitHandler:
        '结束本次调用，释放API资源
        u8apiBroker.Disconnect()

        u8apiBroker = Nothing

        Exit Sub
ErrHandler:
        MsgBox(Err.Description, vbInformation, "提示")
        GoTo ExitHandler
    End Sub

   

    '获取最后一张采购订单号
    Private Sub getLastCode()
        Dim rs As New ADODB.Recordset

        TextBox1.Text = ""
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open("select cPOID from PO_Pomain where POID=(select max(POID) from PO_Pomain)", conn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        If Not rs.EOF Then
            TextBox1.Text = rs.Fields("cPOID").Value
        End If
        rs.Close()
        rs = Nothing
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Char.IsDigit(e.KeyChar) Or e.KeyChar = Chr(8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim c As Integer
        Dim yfhrq As String = Format("2016-6-22  0:00:00", "yyyy-MM-dd")
        MsgBox(yfhrq)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim inv As New Inventory(TextBox2.Text)
        MsgBox(inv.cInvCode + "+++++" + inv.cInvAddCode)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim dh As String = Format(Now(), "yyyy-MM-dd")
        dh = Replace(dh, "-", "")
        MsgBox(dh)

        'Dim x As Array = test()
        'For i = 0 To x.Length - 1
        '    MsgBox(CStr(i) + "||||||" + x(i).ToString)
        'Next

    End Sub
    Private Sub SOImport()
        On Error GoTo ErrHandler
        Dim v As Integer
        'Dim x As Integer = 0 - CInt(TextBox1.Text)


        Dim u8EnvCtx As New U8EnvContext
        u8EnvCtx.U8Login = u8login

        '设置上下文参数
        u8EnvCtx.SetApiContext("VoucherType", 12)  '上下文数据类型：int，含义：单据类型：12

        '第三步：构造ApiBroker对象,调用Connect,传入Api的地址标识(Url)，传入上下文
        Dim u8apiBroker As New U8ApiComBroker

        u8apiBroker.Connect("U8API/SaleOrder/Save", u8EnvCtx)

        '给BO表头参数DomHead赋值，此BO参数的业务类型为采购订单，属表头参数。BO参数均按引用传递
        '提示：给BO表头参数DomHead赋值有两种方法
        '方法一是直接传入MSXML2.DOMDocument对象
        Dim domHead As New MSXML2.DOMDocument   '单据表头DOM
        Dim domBody As New MSXML2.DOMDocument   '单据表体DOM
        Dim eleHead As IXMLDOMElement
        Dim eleBody As IXMLDOMElement
        Dim ele As IXMLDOMElement
        Dim rs As New ADODB.Recordset
        Dim strSQL As String
        Dim strSOID As String

        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        '查询采购订单表头视图，获取表头DOM结构
        '如果有表头扩展自定义项，则可以关联PO_Pomain_extradefine表
        'editprop（单据编辑属性）：A表新增单据，M表修改单据，D表删除单据
        '新增时只需要一个空的DOM结构，所以查询条件为where 1=0
        strSQL = "select * from SaleOrderQ where 1=0"
        rs.Open(strSQL, conn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        rs.Save(domHead, ADODB.PersistFormatEnum.adPersistXML)
        rs.Close()

        '增加表头数据节点z:row
        eleHead = domHead.selectSingleNode("//rs:data")
        ele = domHead.createElement("z:row")
        eleHead.appendChild(ele)

        strSOID = "0"



        setAttribute(ele, "id", "0000000001")   '主关键字段，Integer类型
        setAttribute(ele, "csocode", "0000000001")   '订 单 号，String类型
        '   setAttribute(ele, "ddate", Format(Now(), "yyyy-MM-dd"))   '订单日期，Date类型
        setAttribute(ele, "ddate", "2016-06-22")   '订单日期，Date类型
        setAttribute(ele, "cbustype", "普通销售")   '业务类型，Integer类型
        setAttribute(ele, "cstname", "普通销售")   '销售类型，String类型
        setAttribute(ele, "ccusabbname", "商用车")   '客户简称，String类型
        setAttribute(ele, "ccuscode", "01")   '客户编码，String类型
        setAttribute(ele, "ccusname", "东风商用车有限公司")   '客户名称，String类型
        setAttribute(ele, "cdepname", "市场部")   '销售部门，String类型
        setAttribute(ele, "itaxrate", "17")   '税率，Double类型
        setAttribute(ele, "cexch_name", "人民币")   '币种，String类型
        setAttribute(ele, "cmaker", u8login.cUserName)   '制单人，String类型
        setAttribute(ele, "cstcode", "01")   '销售类型编号，String类型
        setAttribute(ele, "cdepcode", "07")   '部门编码，String类型
        setAttribute(ele, "iexchrate", "1")   '汇率，Double类型
        setAttribute(ele, "cdefine10", "5800")   '到货方，String类型



        u8apiBroker.AssignNormalValue("DomHead", domHead)

        strSQL = "select * from SaleOrderSQ where 1=0"
        rs.Open(strSQL, conn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        rs.Save(domBody, ADODB.PersistFormatEnum.adPersistXML)
        rs.Close()
        rs = Nothing


        '增加表体数据节点z:row
        eleBody = domBody.selectSingleNode("//rs:data")

     
        ele = domBody.createElement("z:row")
        eleBody.appendChild(ele)

     
        Dim yfhrq As String '= Format(CDate(dt2.Rows(j)("sdate").ToString), "yyyy-MM-dd")
   
        yfhrq = "2016-06-23"
    
        setAttribute(ele, "cinvname", "支架总成-空气滤清器")   '存货名称，String类型
        setAttribute(ele, "cinvcode", "1030000216")   '存货编码，String类型
        ' setAttribute(ele,"autoid", "字段值")   '销售订单 2，Integer类型
        setAttribute(ele, "iquantity", "5800")   '数量，Double类型
        setAttribute(ele, "dpredate", yfhrq)   '预发货日期，Date类型
        setAttribute(ele, "dpremodate", yfhrq)   '预完工日期，Date类型
        setAttribute(ele, "id", "0000000001")   '主表id，Integer类型
        'domBody(y).SetValue("iinvexchrate", "字段值")   '换算率，Double类型
        'domBody(y).SetValue("cunitid", "字段值")   '销售单位编码，String类型
        'domBody(y).SetValue("cinva_unit", "字段值")   '销售单位，String类型
        '   setAttribute(ele, "cinvm_unit", inv.cComUnitCode)   '主计量单位，String类型
        'domBody(y).SetValue("igrouptype", "字段值")   '单位类型，Integer类型
        '    setAttribute(ele, "cgroupcode", inv.cGroupCode)   '计量单位组，String类型
        'domBody(y).SetValue("dreleasedate", "字段值")   '预留失效日期，Date类型
        setAttribute(ele, "editprop", "A")   '编辑属性：A表新增，M表修改，D表删除，String类型

        'y += 1
        'v = y

        '    Next

        u8apiBroker.AssignNormalValue("domBody", domBody)

        '给普通参数VoucherState赋值。此参数的数据类型为Integer，此参数按值传递，表示状态:0增加;1修改
        u8apiBroker.AssignNormalValue("VoucherState", 0)  '参数类型：Integer

        '该参数vNewID为INOUT型普通参数。此参数的数据类型为String，此参数按值传递。在API调用返回时，可以通过GetResult("vNewID")获取其值
        u8apiBroker.AssignNormalValue("vNewID", "000000002")  '参数类型：String

        '给普通参数DomConfig赋值。此参数的数据类型为MSXML2.IXMLDOMDocument2，此参数按引用传递，表示ATO,PTO选配
        Dim CurDom As New DOMDocument
        u8apiBroker.AssignNormalValue("DomConfig", CurDom)  '参数类型：MSXML2.IXMLDOMDocument2

        '第五步：调用API
        If u8apiBroker.InvokeApi() = False Then
            MsgBox(u8apiBroker.GetLastError())
            '第六步：错误处理
            If u8apiBroker.ErrorType = ExceptionType.Business Then

                '处理API业务错误
            ElseIf u8apiBroker.ErrorType = ExceptionType.System Then

                '处理系统错误
            End If
        Else
            '第七步：获取返回结果

            '获取返回值
            '获取普通返回值。此返回值数据类型为String，此参数按值传递，表示成功为空串
            Dim result As String
            result = CStr(u8apiBroker.GetReturnValue())
            '获取out/inout参数值

            '获取普通INOUT参数vNewID。此返回值数据类型为String，在使用该参数之前，请判断是否为空
            Dim vNewIDRet As String
            vNewIDRet = CStr(u8apiBroker.GetResult("vNewID"))

        End If
        '结束本次调用，释放API资源
        u8apiBroker.Disconnect()

        u8apiBroker = Nothing

        MsgBox("导入成功", MsgBoxStyle.OkOnly, "提示")

        Exit Sub
ErrHandler:
        '   MsgBox(v)
        MsgBox(Err.Description)
    End Sub
    Public Sub showprogressbar()

        Dim pr As New waitForm
        If pr.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

    End Sub

    Public Sub CallU8Api()
        On Error GoTo ErrHandler

        Dim u8EnvCtx As New U8EnvContext
        u8EnvCtx.U8Login = u8login

        '第三步：构造ApiBroker对象,调用Connect,传入Api的地址标识(Url)，传入上下文
        Dim u8apiBroker As New U8ApiComBroker
        u8apiBroker.Connect("U8API/Forecast/ForecastAdd", u8EnvCtx)

        Dim extbo As ExtensionBusinessEntity
        extbo = u8apiBroker.GetExtBoEntity("extbo")


        '************************************* 主表 **********************************

        '----------------------------------- 必输字段 --------------------------------
        extbo(0).SetValue("ForecastId", "100000001")   '主键，Integer类型
        extbo(0).SetValue("FoCode", "100000001")   '预测单号(必须)，String类型
        extbo(0).SetValue("DocDate", "2016-06-22")   '单据日期(必须)，Date类型
        extbo(0).SetValue("MpsFlag", "2")   '单据类别(必须:1MPS/2MRP)，Integer类型
        extbo(0).SetValue("Version", "V1")   '预测版本号(必须)，String类型

        extbo(0).SetValue("Define_1", "5813")   '表头自定义项1，String类型



        '******************************** 子表[ForecastDetail] ***************************

        Dim ForecastDetail As ExtensionBusinessEntity
        ForecastDetail = extbo(0).GetSubEntity("ForecastDetail")


        '----------------------------------- 必输字段 --------------------------------
        ForecastDetail(0).SetValue("DInvCode", "1010000001")   '物料编码(必须)，String类型
        ForecastDetail(0).SetValue("DStartDate", "2016-06-22")   '起始日期(必须)，Date类型
        ForecastDetail(0).SetValue("DEndDate", "2016-06-22")   '结束日期(必须)，Date类型
        ForecastDetail(0).SetValue("DFQty", "600")   '预测数量(必须)，Double类型
        ForecastDetail(0).SetValue("DAvgType", "0")   '均化类型(必须:0不均化/1日均化/2周均化/3月均化/4时格均化)，Integer类型
        ForecastDetail(0).SetValue("DAvgRounded", "0")   '均化取整(必须:0/不取整/1取上整/2取下整)，Integer类型

        '---------------------------------- 非必输字段 -------------------------------
        'ForecastDetail(0).SetValue("DInvAddCode", "字段值")   '物料代号(导出用)，String类型
        'ForecastDetail(0).SetValue("DInvName", "字段值")   '物料名称(导出用)，String类型
        'ForecastDetail(0).SetValue("DInvStd", "字段值")   '物料规格(导出用)，String类型


        '第五步：调用API
        If u8apiBroker.InvokeApi() = False Then
            '第六步：错误处理
            MsgBox(u8apiBroker.GetLastError())
            If u8apiBroker.ErrorType = ExceptionType.Business Then
                '处理API业务错误
            ElseIf u8apiBroker.ErrorType = ExceptionType.System Then
                '处理系统错误
            End If
        Else
            '第七步：获取返回结果

            '获取返回值
            '获取普通返回值。此返回值数据类型为Boolean，此参数按值传递，表示返回值: true:成功, false: 失败
            Dim result As Boolean
            result = CBool(u8apiBroker.GetReturnValue())
        End If

        '结束本次调用，释放API资源
        u8apiBroker.Disconnect()

        u8apiBroker = Nothing
        Exit Sub
ErrHandler:
        MsgBox(Err.Description)
    End Sub
    Function GetTablename() As String  'i表示第几个sheet，大于0
        'Dim sConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:/123.xls;Extended Properties=Excel 8.0;"
        'Dim cn As OleDbConnection = New OleDbConnection(sConnectionString)

        'cn.Open()

        Dim _Connectstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<FilePath>;Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
        excConn = New OleDb.OleDbConnection(_Connectstring.Replace("<FilePath>", "c:/1.xlsx"))
        excConn.Open()
        Dim tb As DataTable = excConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        For Each row In tb.Rows
            MsgBox(row.Item(2).ToString())
        Next

        Return ""
        '   Return tb.Rows(i - 1)("TABLE_NAME") '第一个


    End Function
    Function test() As Array
        Dim vList As New List(Of String)
        Try
            Dim strConn As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<FilePath>;Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
            Dim conn As OleDbConnection
            conn = New OleDb.OleDbConnection(strConn.Replace("<FilePath>", "c:/1.xlsx"))

            conn.Open()

            Dim sheetNames As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

            conn.Close()
            Dim vName As String = String.Empty
            Dim pOUTPres As New List(Of String)

            For i = 0 To sheetNames.Rows.Count - 1
                If sheetNames.Rows(i)(2).ToString().Trim().Contains("OUTPres") And i > 0 Then
                    If sheetNames.Rows(i)(2).ToString().Trim().Contains(sheetNames.Rows(i - 1)(2).ToString().Trim() + "OUTPres") Then
                        Continue For
                    End If
                End If
                pOUTPres.Add(sheetNames.Rows(i)(2).ToString().Trim())
            Next

            Dim vSheets As String() = pOUTPres.ToArray()
            Dim pSheetName As String = String.Empty



            For i = 0 To vSheets.Length - 1
                Dim pStart As String = vSheets(i).Substring(0, 1)
                Dim pEnd As String = vSheets(i).Substring(vSheets(i).Length - 1, 1)
                If pStart = "'" And pEnd = "'" Then
                    vSheets(i) = vSheets(i).Substring(1, vSheets(i).Length - 2)
                End If

                Dim pChar As Char() = vSheets(i).ToCharArray
                pSheetName = String.Empty
                For j = 0 To pChar.Length - 1
                    If j < pChar.Length - 1 Then
                        If pChar(j).ToString = "'" And pChar(j + 1).ToString = "'" Then
                            pSheetName += pChar(j).ToString
                            j = j + 1
                        Else
                            pSheetName += pChar(j).ToString
                        End If
                    Else
                        pSheetName += pChar(j).ToString
                    End If

                Next
                vSheets(i) = pSheetName
            Next



            For i = 0 To vSheets.Length - 1
                If vList.IndexOf(vSheets(i).ToLower) = -1 Then
                    vList.Add(vSheets(i))
                End If
            Next

            Dim ptList As New List(Of String)
            For j = 0 To vList.Count - 1
                ptList.Add(vList(j))
            Next


            For i = 0 To ptList.Count - 1
                If ptList(i).ToString().Contains("FilterDatabase") Or ptList(i).ToString().Contains("Print_Titles") _
                     Or ptList(i).ToString().Contains("_xlnm#Database") Or ptList(i).ToString().Contains("Print_Area") _
                     Or ptList(i).ToString().Contains("_xlnm.Database") Or ptList(i).ToString().Contains("ExternalData") _
                     Or ptList(i).ToString().Contains("DRUG_IMP_STOCK") Or ptList(i).ToString().Contains("Sheet1$zy") _
                     Or ptList(i).ToString().Contains("Sheet1$xy") Or ptList(i).ToString().Contains("data_xy_zcy") _
                     Or ptList(i).ToString().Contains("Results") Then

                    vList.Remove(ptList(i).ToString)

                End If

            Next

            If vList.Count > 1 Then
                Dim pCheckList As New List(Of String)
                For j = 0 To vList.Count - 1
                    pCheckList.Add(vList(j))
                Next
                conn.Open()
                Dim pComm As New OleDbCommand
                pComm.Connection = conn

                For i = 0 To pCheckList.Count - 1
                    Try
                        pComm.CommandText = String.Format("select count(*) from [{0}] where 1=0", pCheckList(i))
                        pComm.ExecuteNonQuery()
                    Catch ex As Exception
                        If ex.Message.Contains("Microsoft Access 数据库引擎找不到对象") Then
                            vList.Remove(pCheckList(i).ToString)
                        End If

                    End Try
                Next
                conn.Close()
            End If

        Catch ex As Exception

        End Try

        Return vList.ToArray

    End Function
End Class