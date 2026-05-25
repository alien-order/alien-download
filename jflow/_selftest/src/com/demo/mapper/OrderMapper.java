package com.demo.mapper;

import java.util.List;
import org.apache.ibatis.annotations.Mapper;

@Mapper
public interface OrderMapper {
    List<Object> selectAll();
    void insertOrder(Object form);
}
