#include "StdAfx.h"
#include "rlDocReactor.h"
#include "RLDrawBackEdge.h"

CRLDocReactor::CRLDocReactor (const bool autoInitAndRelease) : AcApDocManagerReactor(), mbAutoInitAndRelease(autoInitAndRelease) {
	if ( autoInitAndRelease ) {
		if (acDocManager)
			acDocManager->addReactor(this);
		else
			mbAutoInitAndRelease =false ;
	}
}

CRLDocReactor::~CRLDocReactor () {
	Detach () ;
}

void CRLDocReactor::Attach () {
	Detach () ;
	if ( !mbAutoInitAndRelease ) {
		if ( acDocManager ) {
			acDocManager->addReactor (this) ;
			mbAutoInitAndRelease =true ;
		}
	}
}

void CRLDocReactor::Detach () {
	if ( mbAutoInitAndRelease ) {
		if ( acDocManager ) {
			acDocManager->removeReactor (this) ;
			mbAutoInitAndRelease =false ;
		}
	}
}

AcApDocManager *CRLDocReactor::Subject () const {
	return (acDocManager) ;
}

bool CRLDocReactor::IsAttached () const {
	return (mbAutoInitAndRelease) ;
}

void   CRLDocReactor::documentActivated(AcApDocument* pActivatedDoc)
{
	CRLDrawBackEdge::InitBackEdge();
}

void CRLDocReactor::documentToBeActivated(AcApDocument* pActivatingDoc )
{
	
}

void CRLDocReactor::documentCreated(AcApDocument* pDocCreating)
{

}

void  CRLDocReactor::documentToBeDestroyed(AcApDocument* pDocToDestroy)
{
	
}