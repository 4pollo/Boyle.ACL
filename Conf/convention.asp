<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [惯例配置文件]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

'/* 该文件请不要修改，如果要覆盖惯例配置的值，可在项目配置文件中设定和惯例不符的配置项 */

'/*项目设定*/
C("APP_DEBUG")             = false		'//是否开启调试模式
C("APP_STATUS")            = "debug"	'//应用调试模式状态调试模式开启后有效默认为debug可扩展并自动加载对应的配置文件
C("APP_FILE_CASE")         = false		'//是否检查文件的大小写对Windows平台有效
C("APP_AUTOLOAD_PATH")     = ""			'//自动加载机制的自动搜索路径,注意搜索顺序
C("APP_TAGS_ON")           = true		'//系统标签扩展开关
C("APP_SUB_DOMAIN_DEPLOY") = false		'//是否开启子域名部署
C("APP_SUB_DOMAIN_RULES")  = "array()"	'//子域名部署规则
C("APP_SUB_DOMAIN_DENY")   = "array()"	'//子域名禁用列表
C("APP_GROUP_LIST")        = ""			'//项目分组设定,多个组之间用逗号分隔,例如"Home,Admin"
C("APP_GROUP_MODE")        = 0 			'//分组模式0普通分组1独立分组
C("APP_GROUP_PATH")        = "Modules"	'//分组目录独立分组模式下面有效
C("ACTION_SUFFIX")         = ""			'//操作方法后缀

'/*Cookie设置*/
C("COOKIE_EXPIRE") = 0 		'//Coodie有效期
C("COOKIE_DOMAIN") = ""		'//Cookie有效域名
C("COOKIE_PATH")   = "/"	'//Cookie路径
C("COOKIE_PREFIX") = ""		'//Cookie前缀避免冲突

'/*默认设定*/
C("DEFAULT_M_LAYER")       = "Model"	'//默认的模型层名称
C("DEFAULT_C_LAYER")       = "Action"	'//默认的控制器层名称
C("DEFAULT_APP")           = "@"		'//默认项目名称，@表示当前项目
C("DEFAULT_LANG")          = "zh-cn"	'//默认语言
C("DEFAULT_THEME")         = ""			'//默认模板主题名称
C("DEFAULT_GROUP")         = "Home"		'//默认分组
C("DEFAULT_MODULE")        = "Index"	'//默认模块名称
C("DEFAULT_ACTION")        = "index"	'//默认操作名称
C("DEFAULT_CHARSET")       = "utf-8"	'//默认输出编码
C("DEFAULT_TIMEZONE")      = ""			'//默认时区
C("DEFAULT_AJAX_RETURN")   = ""			'//默认AJAX数据返回格式,可选JSONXML...
C("DEFAULT_JSONP_HANDLER") = "jsonpReturn"		'//默认JSONP格式返回的处理方法
C("DEFAULT_FILTER")        = "htmlspecialchars"	'//默认参数过滤方法用于$this->_get("变量名");$this->_post("变量名")...

'/*数据库设置*/
C("DB_TYPE")="access"'//数据库类型
C("DB_HOST")="localhost"'//服务器地址
C("DB_NAME")=""'//数据库名
C("DB_USER")="root"'//用户名
C("DB_PWD")=""'//密码
C("DB_PORT")=""'//端口
C("DB_PREFIX")="BOYLE_"'//数据库表前缀
C("DB_FIELDTYPE_CHECK")=false'//是否进行字段类型检查
C("DB_FIELDS_CACHE")=true'//启用字段缓存
C("DB_CHARSET")="utf8"'//数据库编码默认采用utf8
C("DB_DEPLOY_TYPE")=0'//数据库部署方式:0集中式(单一服务器),1分布式(主从服务器)
C("DB_RW_SEPARATE")=false'//数据库读写是否分离主从式有效
C("DB_MASTER_NUM")=1'//读写分离后主服务器数量
C("DB_SLAVE_NO")=""'//指定从服务器序号
C("DB_SQL_BUILD_CACHE")=false'//数据库查询的SQL创建缓存
C("DB_SQL_BUILD_QUEUE")="file"'//SQL缓存队列的缓存方式支持filexcache和apc
C("DB_SQL_BUILD_LENGTH")=20'//SQL缓存的队列长度
C("DB_SQL_LOG")=false'//SQL执行日志记录

'/*数据缓存设置*/
C("DATA_CACHE_TIME")=0'//数据缓存有效期0表示永久缓存
C("DATA_CACHE_COMPRESS")=false'//数据缓存是否压缩缓存
C("DATA_CACHE_CHECK")=false'//数据缓存是否校验缓存
C("DATA_CACHE_PREFIX")=""'//缓存前缀
C("DATA_CACHE_TYPE")="File"'//数据缓存类型,支持:File|Db|Apc|Memcache|Shmop|Sqlite|Xcache|Apachenote|Eaccelerator
C("DATA_CACHE_PATH")=""'//缓存路径设置(仅对File方式缓存有效)
C("DATA_CACHE_SUBDIR")=false'//使用子目录缓存(自动根据缓存标识的哈希创建子目录)
C("DATA_PATH_LEVEL")=1'//子目录缓存级别

'/*错误设置*/
C("ERROR_MESSAGE")="页面错误！请稍后再试～"'//错误显示信息,非调试模式有效
C("ERROR_PAGE")=""'//错误定向页面
C("SHOW_ERROR_MSG")=false'//显示错误信息
C("TRACE_EXCEPTION")=false'//TRACE错误信息是否抛异常针对trace方法

'/*日志设置*/
C("LOG_RECORD")=false'//默认不记录日志
C("LOG_TYPE")=3'//日志记录类型0系统1邮件3文件4SAPI默认为文件方式
C("LOG_DEST")=""'//日志记录目标
C("LOG_EXTRA")=""'//日志记录额外信息
C("LOG_LEVEL")="EMERG,ALERT,CRIT,ERR"'//允许记录的日志级别
C("LOG_FILE_SIZE")=2097152'//日志文件大小限制
C("LOG_EXCEPTION_RECORD")=false'//是否记录异常信息日志

'/*SESSION设置*/
C("SESSION_AUTO_START")=true'//是否自动开启Session
C("SESSION_OPTIONS")="array()"'//session配置数组支持typenameidpathexpiredomian等参数
C("SESSION_TYPE")=""'//sessionhander类型默认无需设置除非扩展了sessionhander驱动
C("SESSION_PREFIX")=""'//session前缀
C("VAR_SESSION_ID")="session_id"'//sessionID的提交变量

'/*模板引擎设置*/
C("TMPL_CONTENT_TYPE")="text/html"'//默认模板输出类型
C("TMPL_ACTION_ERROR")="Tpl/dispatch_jump.tpl"'//默认错误跳转对应的模板文件
C("TMPL_ACTION_SUCCESS")="Tpl/dispatch_jump.tpl"'//默认成功跳转对应的模板文件
C("TMPL_EXCEPTION_FILE")="Tpl/think_exception.tpl"'//异常页面的模板文件
C("TMPL_DETECT_THEME")=false'//自动侦测模板主题
C("TMPL_TEMPLATE_SUFFIX")=".html"'//默认模板文件后缀
C("TMPL_FILE_DEPR")="/"'//模板文件MODULE_NAME与ACTION_NAME之间的分割符

'/*URL设置*/
C("URL_CASE_INSENSITIVE")=false'//默认false表示URL区分大小写true则表示不区分大小写
C("URL_MODEL")=1'//URL访问模式,可选参数0、1、2,代表以下三种模式：0(普通模式);1(单参数模式);2(REWRITE模式);
C("URL_PATHINFO_DEPR")="/"'//单参数模式下，各参数之间的分割符号
C("URL_PATHINFO_FETCH")="ORIG_PATH_INFO,REDIRECT_PATH_INFO,REDIRECT_URL"'//用于兼容判断PATH_INFO参数的SERVER替代变量列表
C("URL_HTML_SUFFIX")=""'//URL伪静态后缀设置
C("URL_PARAMS_BIND")=true'//URL变量绑定到Action方法参数
C("URL_404_REDIRECT")=""'//404跳转页面部署模式有效

'/*系统变量名称设置*/
C("VAR_GROUP")="g"'//默认分组获取变量
C("VAR_MODULE")="m"'//默认模块获取变量
C("VAR_ACTION")="a"'//默认操作获取变量
C("VAR_AJAX_SUBMIT")="ajax"'//默认的AJAX提交变量
C("VAR_JSONP_HANDLER")="callback"
C("VAR_PATHINFO")="s"'//单参数模式获取变量例如?s=/module/action/id/1后面的参数取决于URL_PATHINFO_DEPR
C("VAR_URL_PARAMS")="_URL_"'//PATHINFOURL参数变量
C("VAR_TEMPLATE")="t"'//默认模板切换变量
C("VAR_FILTERS")="filter_exp"'//全局系统变量的默认过滤方法多个用逗号分割

C("OUTPUT_ENCODE")=true'//页面压缩输出
C("HTTP_CACHE_CONTROL")="private"'//网页缓存控制

%>