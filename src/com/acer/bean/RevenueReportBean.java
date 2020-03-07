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
public class RevenueReportBean 
{
    String operator_Id ;
    String operator_name ;
    String business_date ;
    String depot_Id ;
    String depot_name ;
    int trx_count;
    int txamt;
    int personal_count;
    int personal_txamt;
    int transfer_MRT_count;
    int transfer_MRT_txamt;
    int transfer_bus_count;
    int transfer_bus_txamt;
    
    // 業者代碼
    public String getOperator_Id()
    {
        return operator_Id;
    }
    
    public void setOperator_Id(String operator_Id)
    {
        this.operator_Id = operator_Id;
    }
    
    
    // 業者名稱
    public String getOperator_name()
    {
        return operator_name;
    }
    
    public void setOperator_name(String operator_name)
    {
        this.operator_name = operator_name;
    }
    
    
    // 營運日期
    public String getBusiness_date()
    {
        return business_date;
    }
    
    public void setBusiness_date(String business_date)
    {
        this.business_date = business_date;
    }
    
    
    // 場站代碼
    public String getDepot_Id()
    {
        return depot_Id;
    }
    
    public void setDepot_Id(String depot_Id)
    {
        this.depot_Id = depot_Id;
    }
    
    
    // 場站名稱
    public String getDepot_name()
    {
        return depot_name;
    }
    
    public void setDepot_name(String depot_name)
    {
        this.depot_name = depot_name;
    }
    
    
    // 卡片扣款筆數
    public int getTrx_count()
    {
        return trx_count;
    }
    
    public void setTrx_count(int trx_count)
    {
        this.trx_count = trx_count;
    }
    
    
    // 卡片扣款金額
    public int getTxamt()
    {
        return txamt;
    }
    
    public void setTxamt(int txamt)
    {
        this.txamt = txamt;
    }
    
    
    // 社福筆數
    public int getPersonal_count()
    {
        return personal_count;
    }
    
    public void setPersonal_count(int personal_count)
    {
        this.personal_count = personal_count;
    }
    
    
    // 社福金額
    public int getPersonal_txamt()
    {
        return personal_txamt;
    }
    
    public void setPersonal_txamt(int personal_txamt)
    {
        this.personal_txamt = personal_txamt;
    }
    
    
    // 捷運轉乘筆數
    public int getTransfer_MRT_count()
    {
        return transfer_MRT_count;
    }
    
    public void setTransfer_MRT_count(int transfer_MRT_count)
    {
        this.transfer_MRT_count = transfer_MRT_count;
    }
    
    
    // 捷運轉乘金額
    public int getTransfer_MRT_txamt()
    {
        return transfer_MRT_txamt;
    }
    
    public void setTransfer_MRT_txamt(int transfer_MRT_txamt)
    {
        this.transfer_MRT_txamt = transfer_MRT_txamt;
    }
    
    
    // 公車轉乘筆數
    public int getTransfer_bus_count()
    {
        return transfer_bus_count;
    }
    
    public void setTransfer_bus_count(int transfer_bus_count)
    {
        this.transfer_bus_count = transfer_bus_count;
    }
    
    
    // 公車轉乘金額
    public int getTransfer_bus_txamt()
    {
        return transfer_bus_txamt;
    }
    
    public void setTransfer_bus_txamt(int transfer_bus_txamt)
    {
        this.transfer_bus_txamt = transfer_bus_txamt;
    }
   
    
}
