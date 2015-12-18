/********************************************************************
	
	purpose:	ÎÄµµ·´Ó¦Æ÷
	Ð»ÕÙÐ÷
*********************************************************************/
#pragma once
#include "aced.h"
#include "acdocman.h"

class CRLDocReactor :public AcApDocManagerReactor
{
protected:
	//----- Auto initialization and release flag.
	bool mbAutoInitAndRelease ;

public:
	CRLDocReactor (const bool autoInitAndRelease =true) ;
	virtual ~CRLDocReactor () ;

	virtual void Attach () ;
	virtual void Detach () ;
	virtual AcApDocManager *Subject () const ;
	virtual bool IsAttached () const ;

public:
	virtual void   documentToBeActivated(AcApDocument* pActivatingDoc );
	virtual void  documentCreated(AcApDocument* pDocCreating);
    virtual void   documentToBeDestroyed(AcApDocument* pDocToDestroy);
	 virtual void   documentActivated(AcApDocument* pActivatedDoc);
};

