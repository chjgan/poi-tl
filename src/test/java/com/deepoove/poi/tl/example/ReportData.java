package com.deepoove.poi.tl.example;

import com.deepoove.poi.data.DocxRenderData;

public class ReportData {
    private String createTime;
    private String dataTime;
    private String titles;
    private String names;
    private DocxRenderData model;

    public String getCreateTime() {
        return createTime;
    }

    public void setCreateTime(String createTime) {
        this.createTime = createTime;
    }

    public String getDataTime() {
        return dataTime;
    }

    public void setDataTime(String dataTime) {
        this.dataTime = dataTime;
    }

    public String getTitles() {
        return titles;
    }

    public void setTitles(String titles) {
        this.titles = titles;
    }

    public String getNames() {
        return names;
    }

    public void setNames(String names) {
        this.names = names;
    }

    public DocxRenderData getModel() {
        return model;
    }

    public void setModel(DocxRenderData model) {
        this.model = model;
    }
}
