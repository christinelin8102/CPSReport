/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package com.acer.bean;

/**
 *
 * @author Christine
 */
public class UnclearedReportBean 
{
    
    String operating_id;
    String operating_name;
    String cal_date;
    String unit_id;
    int txamt_all;
    int personal_txamt_all;
    int transfer_txamt_all;
    int uncleared_count_c;
    int txamt_c;
    int personal_txamt_c;
    int transfer_txamt_c;
    int uncleared_count_e;
    int txamt_e;
    int personal_txamt_e;
    int transfer_txamt_e;
    String update_time;
    
    // 業者代碼
    public String getOperating_id()
    {
        return operating_id;
    }
    
    public void setOperating_id(String operating_id)
    {
        this.operating_id = operating_id;
    }
    
    
    // 業者名稱
    public String getOperating_name()
    {
        return operating_name;
    }
    
    public void setOperating_name(String operating_name)
    {
        this.operating_name = operating_name;
    }
    
    
    // 清分日期
    public String getCal_date()
    {
        return cal_date;
    }
    
    public void setCal_date(String cal_date)
    {
        this.cal_date = cal_date;
    }
    
    // 票證公司代碼
    public String getUnit_id()
    {
        return unit_id;
    }
    
    public void setUnit_id(String unit_id)
    {
        this.unit_id = unit_id;
    }

    
    // 卡片扣款金額(CPS全部)
    public int getTxamt_all()
    {
        return txamt_all;
    }
    
    public void setTxamt_all(int txamt_all)
    {
        this.txamt_all = txamt_all;
    }
    
    // 社福扣款金額(CPS全部)
    public int getPersonal_txamt_all()
    {
        return personal_txamt_all;
    }
    
    public void setPersonal_txamt_all(int personal_txamt_all)
    {
        this.personal_txamt_all = personal_txamt_all;
    }
    
    // 轉乘扣款金額(CPS全部)
    public int getTransfer_txamt_all()
    {
        return transfer_txamt_all;
    }
    
    public void setTransfer_txamt_all(int transfer_txamt_all)
    {
        this.transfer_txamt_all = transfer_txamt_all;
    }
    
    // 已上傳未清分筆數
    public int getUncleared_count_c()
    {
        return uncleared_count_c;
    }
    
    public void setUncleared_count_c(int uncleared_count_c)
    {
        this.uncleared_count_c = uncleared_count_c;
    }
    
    // 卡片扣款金額(已上傳未清分)
    public int getTxamt_c()
    {
        return txamt_c;
    }
    
    public void setTxamt_c(int txamt_c)
    {
        this.txamt_c = txamt_c;
    }
    
    // 社福扣款金額(已上傳未清分)
    public int getPersonal_txamt_c()
    {
        return personal_txamt_c;
    }
    
    public void setPersonal_txamt_c(int personal_txamt_c)
    {
        this.personal_txamt_c = personal_txamt_c;
    }
    
    // 轉乘扣款金額(已上傳未清分)
    public int getTransfer_txamt_c()
    {
        return transfer_txamt_c;
    }
    
    public void setTransfer_txamt_c(int transfer_txamt_c)
    {
        this.transfer_txamt_c = transfer_txamt_c;
    }
    
    // 清分錯誤筆數
    public int getUncleared_count_e()
    {
        return uncleared_count_e;
    }
    
    public void setUncleared_count_e(int uncleared_count_e)
    {
        this.uncleared_count_e = uncleared_count_e;
    }
    
    // 卡片扣款金額(清分錯誤)
    public int getTxamt_e()
    {
        return txamt_e;
    }
    
    public void setTxamt_e(int txamt_e)
    {
        this.txamt_e = txamt_e;
    }
    
    // 社福扣款金額(清分錯誤)
    public int getPersonal_txamt_e()
    {
        return personal_txamt_e;
    }
    
    public void setPersonal_txamt_e(int personal_txamt_e)
    {
        this.personal_txamt_e = personal_txamt_e;
    }
    
    // 轉乘扣款金額(清分錯誤)
    public int getTransfer_txamt_e()
    {
        return transfer_txamt_e;
    }
    
    public void setTransfer_txamt_e(int transfer_txamt_e)
    {
        this.transfer_txamt_e = transfer_txamt_e;
    }
    
    // 更新時間
    public String getUpdate_time()
    {
        return update_time;
    }
    
    public void setUpdate_time(String update_time)
    {
        this.update_time = update_time;
    }
    

    
}
