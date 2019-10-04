package io.github.learnapachepoixssf.model;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Table;
import javax.persistence.Transient;

import org.springframework.data.domain.Persistable;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Entity
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Table(name = "WIDGET")
public class Widget implements Persistable<Long> {

	@Id
	@GeneratedValue(strategy = GenerationType.SEQUENCE, generator = "SEQ_WIDGET")
	@Column(name = "ID")
	private Long id;

	@Column(name = "NAME", columnDefinition = "VARCHAR(50)", nullable = false, length = 50)
	private String name;

	@Transient
	public boolean isNew() {
		return getId() == null;
	}
}
