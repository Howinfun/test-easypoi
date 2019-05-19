package com.hyf.teseasypoi.entity;

import lombok.Data;

import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Table;
import java.io.Serializable;

/**
 * @author howinfun
 * @version 1.0
 * @desc
 * @date 2019/5/18
 * @company WCWC
 */
@Data
@Table(name="clazz")
public class Clazz implements Serializable {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Integer id;

    private String name;
}
