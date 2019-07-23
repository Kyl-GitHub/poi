# POI 导出Excel详解 #

## 1、定义导出Excel的工具类 ##

    <!-- poi office -->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>${poi.version}</version>
        </dependency>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>${poi.version}</version>
        </dependency>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml-schemas</artifactId>
            <version>${poi.version}</version>
        </dependency>

**通过构造方法初始化导出的表格对象的公共内容，例：表格标题，表格表头等**

    public ExportExcel(String title, Class<?> cls){
		this(title, cls, 1);
	}
	这个构造方法中，title为Excel中的标题，cls为要导出的数据对应的对象
	
	public ExportExcel(String title, Class<?> cls, int type){
		...
	}
	这个构造方法为具体实现，其中type为导出类型（1:导出数据；2：导出模板），你也可以定义更多其他想要导出的类型
**在构造方法中初始化导出对象**

    Field[] fs = cls.getDeclaredFields(); 通过java反射获取到对象属性
	ExcelField ef = field.getAnnotation(ExcelField.class); 遍历对象属性，通过反射获取该属性是否添加了ExcelField注解，如果有该注解，则表示这个属性需要导出到Excel中，于是将该属性保存到内存中annotationList.add(new Object[]{ef, field});

	Method[] ms = cls.getDeclaredMethods(); 通过java反射获取到对象中的方法
	ExcelField ef = method.getAnnotation(ExcelField.class); 遍历对象方法，通过反射获取方法上是否添加了ExcelField注解，如果有该注解，则表示这个方法对应的属性需要导出到Excel中，于是将该属性保存到内存中annotationList.add(new Object[]{ef, method});
**遍历存入的带注解的属性**
    创建表头list，List<String> headerList = Lists.newArrayList();
	
	遍历annotationList ，将对象属性放入headerList，初始化要导出的Excel对象 initialize(title, headerList);该方法中title为Excel标题，headerList为表头
**初始化**
	private void initialize(String title, List<String> headerList) {...}
	该方法中title为标题，headerList为表头
	
	new SXSSFWorkbook(500); 创建工作表，设置缓存记录数为500，该值默认为100
	
	wb.createSheet("sheet"); 通过工作表对象创建sheet，命名为sheet

	createStyles(wb); 为工作表创建样式
	
	Row titleRow = sheet.createRow(rownum++); 为标题创建一行，rownum为行号，默认为0，rownum++表示工作表中的第一行

	titleRow.setHeightInPoints(30); 设置行高，30为像素值，和Excel中设置的行高一致

	Cell titleCell = titleRow.createCell(0); //创建一个单元格
	titleCell.setCellStyle(styles.get("title"));//给单元格设置样式
	titleCell.setCellValue(title);//给单元格设置值

	sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(),titleRow.getRowNum(), titleRow.getRowNum(), headerList.size()-1)); //合并单元格，标题为一行，一个单元格，其宽度与表格一样。 单元格的范围有四个参数确定，分别为(起始行号，终止行号， 起始列号，终止列号） 这个单元格中，起始行为1，终止行为1，起始列为1，终止列为表头的长度

	Row headerRow = sheet.createRow(rownum++); //创建第二行，该行为表头
	headerRow.setHeightInPoints(20); //设置行高为20
	Cell cell = headerRow.createCell(i); 遍历headerList，创建单元格
	cell.setCellValue(headerList.get(i)); 设置单元格的值 
	
	sheet.autoSizeColumn(i); //标识该字段对应的单元格自适应宽度



----------
以上为Excel初始化的过程，这期间你可以给单元格设置漂亮的样式

## 2、设置导出的数据 ##

    exportExcel.setDataList(list); //通过上述初始化的Excel对象，设置当前要导出的数据

	public <E> ExportExcel setDataList(List<E> list){...} // 在该方法中初始化数据 

	sheet.createRow(rownum++); 创建行，rownum以上述初始化导出Excel对象为基础，rownum 从3开始 创建第三行，设置数据

	ExcelField ef = (ExcelField)os[0]; //遍历annotationList ，对有注解的方法或者属性进行转换
	
	val = Reflections.invokeGetter(e, ef.value()); 通过java反射，获取到实体的值
	
	public Cell addCell(Row row, int column, Object val, int align, Class<?> fieldType){...} // 添加单元格，row表示添加的行，column 添加列号，val值，align 对其方式，fieldType为导出Excel对象。

	

----------
以上为设置Excel导出数据的过程，该过程中的数据为数据库中的原数据，其中每个字段的值与初始化表头是对应的属性值是一致的。

## 3、导出数据 ##

	public ExportExcel write(HttpServletResponse response, String fileName) throws IOException{...} 通过response对象获取到数据流对象，用流将文件写出到用户主机。 fileName 为输出的文件名

	response.reset();
    response.setContentType("application/octet-stream; charset=utf-8");
    response.setHeader("Content-Disposition", "attachment; filename="+Encodes.urlEncode(fileName));
	write(response.getOutputStream());
	
	


	
