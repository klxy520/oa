package com.xnky.oa.util.constant;

/**
 * 
 * ContentUtil
 *
 * @author: 杨永川
 * @version: 1.0, 2015年10月26日
 */

public class ContentUtils {
    public static final String  VERSION                          = "/v1";
    public static final String  CHARACTER_SET_UTF_8              = "UTF-8";
    public static final String  CONTENT_TYPE                     = "text/xml; charset=utf-8";
    public static final String  TRUE                             = "true";
    public static final String  FALSE                            = "false";
    public static final String  SUCCESS                          = "success";
    public static final String  ERROR                            = "error";
    public static final String  FAILED                           = "failed";
    public static final String  USER                             = "user";
    public static final String  SESSION_USER                     = "session_user";
    public static final String  SESSION_USER_NAME                = "session_user_name";
    public static final String  COMMA                            = ",";
    public static final String  TIME_FORMAT1                     = "yyyy-MM-dd HH:mm:ss";
    public static final String  TIME_FORMAT2                     = "yyyy.MM.dd HH:mm";
    public static final String  TIME_FORMAT3                     = "yyyyMMddHHmmss";
    public static final String  TIME_FORMAT_SIMPLE               = "yyyy-MM-dd";
    public static final String  TIME_FORMAT_MONTH                = "yyyy-MM";
    public static final String  TIME_FORMAT_YEAR                 = "yyyy";

    public static final String  CONTENTTYPE_JSON                 = "application/json";
    public static final String  HEADER_ADVERTISEMENT             = "headerAdvertisement";
    public static final String  VERTICAL                         = "\\|";
    public static final String  TWO_BLANK                        = "  ";
    public static final String  COLON                            = ":";
    public static final String  PLUS_SIGN                        = "\\+";
    public static final String  POINT_SIGN                       = ".";
    public static final String  YYYY_MM_DD                       = "yyyy-MM-dd";
    public static final String  UNDERLINE                        = "_";

    public static final String  HTTP                             = "http://";
    public static final String  LINE_FEED                        = "\r\n";
    public static final String  PAGEINDEX                        = "pageIndex";
    public static final String  PAGESIZE                         = "pageSize";

    public static final String  APPID                            = "appid";
    public static final String  APPSECRET                        = "appsecret";
    public static final String  CYCLE_ACTIVITY_ID                = "cycleActivityId";
    public static final String  ACTIVITY_NOT_START               = "activity_not_start";
    public static final String  SERVICE_ERROR                    = "service_error";
    public static final String  JOIN_SUCCESS                     = "join_success";
    public static final String  JOIN_FAILED                      = "join_failed";
    public static final String  OPENID                           = "openid";
    public static final String  REMAINING_LOTTERY_NUMBER_IS_NULL = "remaining_lottery_number_is_null";
    public static final String  DOMAIN                           = "domain";
    public static final String  CURRENT_SHARE_URL                = "current_share_url";
    public static final String  SHARE_TITLE                      = "share_title";
    public static final String  SHARE_DESC                       = "share_desc";

    public static final String  SHARE_LINK                       = "share_link";
    public static final String  SHARE_IMGURL                     = "share_imgurl";

    public static final String  MENU_ROOT_ID                     = "menu_root_id";

    public static final String  FILE_UPLOAD_ROOT_PATH            = "file_upload_root_path";

    // 登录失败 security异常
    public static final String  SPRING_SECURITY_LAST_EXCEPTION   = "SPRING_SECURITY_LAST_EXCEPTION";
    // 商品上架状态
    public static final String  GOODS_STATE_ONSHELVES            = "1";
    // 商品下架状态
    public static final String  GOODS_STATE_OFFSHELVES           = "0";

    // 条件查询 字段前缀标记
    public static final String  FILTER_QUERY_FLAG                = "filter_";
    // 数据库字段状态 0 启用
    public static final Integer DB_FIELD_STATUS_ENABLE           = 0;
    // 数据库字段状态 0 禁用
    public static final Integer DB_FIELD_STATUS_DISABLE          = 1;

    // 密码加密前面加特定的字符串
    public static final String  USER_PWD_ENCP_PREFIX             = "OPIJHBNIYTFGAWSDFGYBHUJK";

    // 登录用户所属平台
    public static final String  LOGIN_USER_PLATFORM              = "login_user_platform";

    // 数据库字段状态 0 未删除
    public static final Integer DB_FIELD_NOT_DELETE              = 0;
    // 数据库字段状态 1 删除
    public static final Integer DB_FIELD_DELETE                  = 1;

    // 积分类型,0.签到
    public static final Integer Integra_Type_Sign                = 0;
    // 积分类型 1.累计奖励
    public static final Integer Integra_Type_Cumulative_Rewards  = 1;
    // 积分类型 2.下单消费
    public static final Integer Integra_Type_Order_Pay           = 2;
    // 积分类型 3.活动奖励
    public static final Integer Integra_Type_Activity_Rewards    = 3;
    // 积分类型 4.退货
    public static final Integer Integra_Type_Order_Return        = 4;
    // 积分类型 5.互动活动消费
    public static final Integer Integra_Type_Activity_Consume    = 5;

    // 积分改变类型(积分增减): 1.增加积分
    public static final Integer integral_regulation_add          = 1;
    // 积分改变类型(积分增减): 0.积分减少
    public static final Integer integral_regulation_reduce       = 0;

    // 互动活动token获取状态: 1:token获取失败
    public static final int     token_retCode_error              = 0;
    // 互动活动token获取状态: 0:token获取成功
    public static final int     token_retCode_success            = 1;
}
