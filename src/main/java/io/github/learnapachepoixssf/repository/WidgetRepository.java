package io.github.learnapachepoixssf.repository;

import java.util.stream.Stream;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import io.github.learnapachepoixssf.model.Widget;

@Repository
public interface WidgetRepository extends JpaRepository<Widget, Long> {

	// Really? Query the entire table?
	@Query("FROM Widget w")
	Stream<Widget> streamAll();

	@Modifying
	@Query(value = "TRUNCATE TABLE Widget", nativeQuery = true)
	void truncateWidgets();
}
