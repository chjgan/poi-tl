package com.deepoove.poi.tl.example;

import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.PictureRenderData;

public class ModelData {
    private String title;
    private PictureRenderData picture;
    private MiniTableRenderData table;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public PictureRenderData getPicture() {
        return picture;
    }

    public void setPicture(PictureRenderData picture) {
        this.picture = picture;
    }

    public MiniTableRenderData getTable() {
        return table;
    }

    public void setTable(MiniTableRenderData table) {
        this.table = table;
    }
}
