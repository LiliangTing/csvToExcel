package com.lianglianglee;

import java.util.List;

/**
 * 爬取到的数据实体.
 */
public class DataEntity {
  private String title;
  private List<String> value;

  public String getTitle() {
    return title;
  }

  public void setTitle(String title) {
    this.title = title;
  }

  public List<String> getValue() {
    return value;
  }

  public void setValue(List<String> value) {
    this.value = value;
  }
}
