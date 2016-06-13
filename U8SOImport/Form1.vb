Imports UFIDA.U8.MomServiceCommon
Imports UFIDA.U8.U8MOMAPIFramework
Imports UFIDA.U8.U8APIFramework
Imports UFIDA.U8.U8APIFramework.Meta
Imports UFIDA.U8.U8APIFramework.Parameter
Imports MSXML2
Imports System.Data
Imports System.Data.OleDb
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
        Dim a(3) As String
        a(0) = "0"
        a(1) = "1"
        a(2) = "2"
        a(3) = "3"
        MsgBox(a(3))
    End Sub
End Class