/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package com.acer.service;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author Christine
 */
public abstract class AbstractRevReport 
{
    //    private static Logger log = Logger.getLogger(AbstractReport.class);
    
    private Connection conn=null;
    private PreparedStatement ps=null;
    private ResultSet rs=null;
//    private String sql=null;
    
    public Connection getConn() {
        return conn;
    }

    /**
     * @param conn the conn to set
     */
    public void setConn(Connection conn) {
        this.conn = conn;
    }
         
    public List<Object> query(String sql)
    {   
        List<Object> list = new ArrayList<Object>();       
        Object obj = null;
        try
        {      
            ps=conn.prepareStatement(sql);
            rs=ps.executeQuery();
        
            while(rs.next())
            {
                obj= rowMapper(rs);
                list.add(obj);
            }
            
            
           
        }
        catch (SQLException connectProblem) 
        {
            System.out.println("doesn't connect! The problem is: "+connectProblem);
        }
        finally
        {
            if( rs != null )
            {
                try {
                    rs.close();
                } catch (SQLException resultSetProblem) {
                    System.out.println("The resultSet doesn't close! The problem is: "+resultSetProblem);
                }
            }
            
            if( ps != null )
            {
                try {
                    ps.close();
                } catch (SQLException preparedStatementProblem) {
                    System.out.println("The PreparedStatement doesn't close! The problem is: "+preparedStatementProblem);
                }
            }
        }  
        return list;
    }
    abstract public Object rowMapper(ResultSet rs) throws SQLException;
}
