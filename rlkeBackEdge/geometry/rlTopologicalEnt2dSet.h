#pragma once
#include "rlTopologicalEntSet.h"

class rlTopologicalEntity2d;
class rlEdge2d;
class rlVertex2d;
class rlTopologicalEnt2dSet :
	public rlTopologicalEntSet
{
public:
	rlTopologicalEnt2dSet();
	~rlTopologicalEnt2dSet();

	template<class T>
	T *_new()
	{
		return __new<T>();
	}
	void _delete(rlTopologicalEntity2d *entity);

	std::map<rlId, rlEdge2d *> edges() const;
	std::map<rlId, rlVertex2d *> vertices() const;
};

