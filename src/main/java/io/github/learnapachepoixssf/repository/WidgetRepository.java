package io.github.learnapachepoixssf.repository;

import java.util.stream.Stream;

import org.springframework.data.jpa.domain.Specification;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import io.github.learnapachepoixssf.model.Widget;

@Repository
public interface WidgetRepository extends JpaRepository<Widget, Long> {

	Stream<Widget> findAll(Specification<Widget> specification);
	
	@Modifying
	@Query(value = "TRUNCATE TABLE Widget", nativeQuery = true)
	void truncateWidgets();
}
